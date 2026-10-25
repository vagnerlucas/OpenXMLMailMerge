using System;
using System.Buffers.Binary;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Reads the pixel size of an image from its header, without decoding it.
    /// Supports PNG, JPEG, GIF, BMP and TIFF, plus EMF and placeable WMF. EMF and WMF are vector formats: their header
    /// gives a physical size, converted to pixels at 96 DPI. A non-placeable WMF has no physical size, so it isn't read.
    /// </summary>
    internal static class ImagePixelSize
    {
        /// <summary>
        /// Reads the pixel size of an image.
        /// </summary>
        /// <param name="data">Byte array of the image data.</param>
        /// <param name="width">Width in pixels.</param>
        /// <param name="height">Height in pixels.</param>
        /// <returns>True when the format is supported and the size is valid.</returns>
        internal static bool TryRead(byte[] data, out int width, out int height)
        {
            width = height = 0;
            if (data == null)
                return false;

            if (IsPng(data))
            {
                width = (int)ReadUInt32BigEndian(data, 16);
                height = (int)ReadUInt32BigEndian(data, 20);
            }
            else if (IsGif(data))
            {
                width = data[6] | data[7] << 8;
                height = data[8] | data[9] << 8;
            }
            else if (IsBmp(data))
            {
                // The DIB header starts at offset 14 with its own size. OS/2 BITMAPCOREHEADER (12 bytes) stores
                // 16-bit dimensions; every later header (BITMAPINFOHEADER and up) stores 32-bit ones.
                if (ReadInt32LittleEndian(data, 14) == 12)
                {
                    width = data[18] | data[19] << 8;
                    height = data[20] | data[21] << 8;
                }
                else
                {
                    width = ReadInt32LittleEndian(data, 18);
                    // A negative height marks a top-down bitmap.
                    height = System.Math.Abs(ReadInt32LittleEndian(data, 22));
                }
            }
            else if (IsTiff(data))
            {
                if (!TryReadTiff(data, out width, out height))
                    return false;
            }
            else if (IsEmf(data))
            {
                // rclFrame: the picture's bounds in 0.01 mm, which applications use as its size.
                width = HundredthsOfMillimeterToPixels(ReadInt32LittleEndian(data, 32) - ReadInt32LittleEndian(data, 24));
                height = HundredthsOfMillimeterToPixels(ReadInt32LittleEndian(data, 36) - ReadInt32LittleEndian(data, 28));
            }
            else if (IsPlaceableWmf(data))
            {
                // Bounding box in logical units, and how many of those units make an inch.
                var unitsPerInch = BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(14));
                if (unitsPerInch == 0)
                    return false;

                width = UnitsToPixels(BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(10)) - BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(6)), unitsPerInch);
                height = UnitsToPixels(BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(12)) - BinaryPrimitives.ReadInt16LittleEndian(data.AsSpan(8)), unitsPerInch);
            }
            else if (!TryReadJpeg(data, out width, out height))
            {
                return false;
            }

            return width > 0 && height > 0;
        }

        /// <summary>
        /// Little-endian ("II*\0") or big-endian ("MM\0*") TIFF header.
        /// </summary>
        internal static bool IsTiff(byte[] data)
        {
            return data.Length >= 8
                   && (data[0] == 'I' && data[1] == 'I' && data[2] == 0x2A && data[3] == 0x00
                       || data[0] == 'M' && data[1] == 'M' && data[2] == 0x00 && data[3] == 0x2A);
        }

        /// <summary>
        /// EMR_HEADER record (type 1) with the " EMF" signature at offset 40.
        /// </summary>
        internal static bool IsEmf(byte[] data)
        {
            return data.Length >= 44
                   && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00
                   && data[40] == ' ' && data[41] == 'E' && data[42] == 'M' && data[43] == 'F';
        }

        /// <summary>
        /// Placeable WMF: the 0x9AC6CDD7 key, then a bounding box and its units per inch.
        /// </summary>
        internal static bool IsPlaceableWmf(byte[] data)
        {
            return data.Length >= 22 && data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A;
        }

        /// <summary>
        /// WMF without the placeable header: a META_HEADER of type 1 (memory) or 2 (disk), 9 words long,
        /// version 0x0100 or 0x0300.
        /// </summary>
        internal static bool IsNonPlaceableWmf(byte[] data)
        {
            return data.Length >= 18
                   && (data[0] == 0x01 || data[0] == 0x02) && data[1] == 0x00
                   && data[2] == 0x09 && data[3] == 0x00
                   && data[4] == 0x00 && (data[5] == 0x01 || data[5] == 0x03);
        }

        /// <summary>
        /// Reads ImageWidth (tag 256) and ImageLength (tag 257) from the first image file directory (IFD).
        /// </summary>
        private static bool TryReadTiff(byte[] data, out int width, out int height)
        {
            width = height = 0;
            var bigEndian = data[0] == 'M';

            var directory = ReadUInt32(data, 4, bigEndian);
            if (directory > data.Length - 2)
                return false;

            var offset = (int)directory;
            var entries = ReadUInt16(data, offset, bigEndian);
            for (var i = 0; i < entries; i++)
            {
                var entry = offset + 2 + i * 12;
                if (entry + 12 > data.Length)
                    return false;

                var tag = ReadUInt16(data, entry, bigEndian);
                if (tag != 256 && tag != 257)
                    continue;

                // The value fits in the entry: a SHORT (type 3) in its first 2 bytes, a LONG (type 4) in all 4.
                var value = ReadUInt16(data, entry + 2, bigEndian) switch
                {
                    3 => ReadUInt16(data, entry + 8, bigEndian),
                    4 => (int)Math.Min(ReadUInt32(data, entry + 8, bigEndian), int.MaxValue),
                    _ => 0
                };

                if (tag == 256)
                    width = value;
                else
                    height = value;
            }

            return width > 0 && height > 0;
        }

        private static int HundredthsOfMillimeterToPixels(int hundredths)
        {
            return (int)Math.Round(Math.Abs(hundredths) * 96.0 / 2540);
        }

        private static int UnitsToPixels(int units, int unitsPerInch)
        {
            return (int)Math.Round(Math.Abs(units) * 96.0 / unitsPerInch);
        }

        private static int ReadUInt16(byte[] data, int offset, bool bigEndian)
        {
            var span = data.AsSpan(offset, 2);
            return bigEndian ? BinaryPrimitives.ReadUInt16BigEndian(span) : BinaryPrimitives.ReadUInt16LittleEndian(span);
        }

        private static uint ReadUInt32(byte[] data, int offset, bool bigEndian)
        {
            var span = data.AsSpan(offset, 4);
            return bigEndian ? BinaryPrimitives.ReadUInt32BigEndian(span) : BinaryPrimitives.ReadUInt32LittleEndian(span);
        }

        internal static bool IsPng(byte[] data)
        {
            return data.Length >= 24
                   && data[0] == 0x89 && data[1] == 'P' && data[2] == 'N' && data[3] == 'G'
                   && data[12] == 'I' && data[13] == 'H' && data[14] == 'D' && data[15] == 'R';
        }

        internal static bool IsGif(byte[] data)
        {
            return data.Length >= 10 && data[0] == 'G' && data[1] == 'I' && data[2] == 'F';
        }

        internal static bool IsBmp(byte[] data)
        {
            return data.Length >= 26 && data[0] == 'B' && data[1] == 'M';
        }

        /// <summary>
        /// Walks the JPEG segments up to the first start-of-frame (SOFn) marker, which holds the size.
        /// </summary>
        private static bool TryReadJpeg(byte[] data, out int width, out int height)
        {
            width = height = 0;
            if (data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8)
                return false;

            var i = 2;
            while (i + 9 < data.Length)
            {
                if (data[i] != 0xFF)
                    return false;

                var marker = data[i + 1];
                switch (marker)
                {
                    case 0xFF:
                        // Fill byte before a marker.
                        i++;
                        continue;
                    // Markers without a length field.
                    case 0x01:
                    case >= 0xD0 and <= 0xD7:
                        i += 2;
                        continue;
                    // SOF0-SOF15, except DHT (C4), JPG (C8) and DAC (CC).
                    case >= 0xC0 and <= 0xCF when marker != 0xC4 && marker != 0xC8 && marker != 0xCC:
                        height = data[i + 5] << 8 | data[i + 6];
                        width = data[i + 7] << 8 | data[i + 8];
                        return true;
                }

                var length = data[i + 2] << 8 | data[i + 3];
                if (length < 2)
                    return false;

                i += 2 + length;
            }

            return false;
        }

        private static uint ReadUInt32BigEndian(byte[] data, int offset)
        {
            return (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);
        }

        private static int ReadInt32LittleEndian(byte[] data, int offset)
        {
            return data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24;
        }
    }
}
