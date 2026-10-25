using System.Collections.Generic;
using System.Globalization;
using OpenXmlMailMerge.Fields;

namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Image element definitions.
    /// </summary>
    internal class ImageElementConfig
    {
        /// <summary>
        /// Size in pixels used when the image's own size can't be read.
        /// </summary>
        public const int DefaultSize = 42;

        /// <summary>
        /// Width in pixels. 0 when not set.
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// Height in pixels. 0 when not set.
        /// </summary>
        public int Height { get; set; }
        /// <summary>
        /// Maximum width in pixels for an image shown at its own size. 0 for no limit.
        /// </summary>
        public int MaxWidth { get; set; }

        /// <summary>
        /// Fills in the dimensions that were not set, from the image's own pixel size.
        /// With neither dimension set, the image keeps its own size, scaled down to <see cref="MaxWidth"/>.
        /// With one dimension set, the other follows the image's aspect ratio.
        /// When the image's size can't be read, a missing dimension falls back to <see cref="DefaultSize"/>.
        /// </summary>
        /// <param name="data">Byte array of the image data.</param>
        public void ResolveSize(byte[] data)
        {
            if (Width > 0 && Height > 0)
                return;

            if (!ImagePixelSize.TryRead(data, out var pixelWidth, out var pixelHeight))
            {
                if (Width <= 0)
                    Width = DefaultSize;
                if (Height <= 0)
                    Height = DefaultSize;
                return;
            }

            if (Width <= 0 && Height <= 0)
            {
                Width = pixelWidth;
                Height = pixelHeight;

                if (MaxWidth > 0 && Width > MaxWidth)
                {
                    Height = Scale(Height, MaxWidth, Width);
                    Width = MaxWidth;
                }
            }
            else if (Width <= 0)
            {
                Width = Scale(pixelWidth, Height, pixelHeight);
            }
            else
            {
                Height = Scale(pixelHeight, Width, pixelWidth);
            }
        }

        private static int Scale(int value, int numerator, int denominator)
        {
            return System.Math.Max(1, (int)System.Math.Round((double)value * numerator / denominator));
        }

        /// <summary>
        /// Reads WIDTH and HEIGHT (in pixels) from the parameters of a field name, for example "IMG@WIDTH=120@HEIGHT=80".
        /// A missing or invalid value leaves that dimension unset.
        /// </summary>
        /// <param name="fieldName">The field name with its parameters.</param>
        public void ParseConfig(string fieldName)
        {
            var parameters = FieldParameters.Parse(fieldName);

            if (TryGetPixels(parameters, "WIDTH", out var width))
                Width = width;

            if (TryGetPixels(parameters, "HEIGHT", out var height))
                Height = height;
        }

        private static bool TryGetPixels(IReadOnlyDictionary<string, string> parameters, string key, out int pixels)
        {
            pixels = 0;
            return parameters.TryGetValue(key, out var value)
                   && int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out pixels);
        }
    }
}
