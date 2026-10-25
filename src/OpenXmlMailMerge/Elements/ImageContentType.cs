namespace OpenXmlMailMerge.Elements
{
    /// <summary>
    /// Detects the content type of an image from its signature bytes.
    /// Supports the raster and vector formats Word displays: PNG, JPEG, GIF, BMP, TIFF, EMF and WMF.
    /// </summary>
    internal static class ImageContentType
    {
        /// <summary>
        /// Content type used when the format is not recognized.
        /// </summary>
        private const string Fallback = "image/jpeg";

        /// <summary>
        /// Detects the content type of image.
        /// </summary>
        /// <param name="data">Byte array of the image data.</param>
        /// <returns>The content type, or <see cref="Fallback"/> when the format is not recognized.</returns>
        internal static string Detect(byte[] data)
        {
            if (data == null)
                return Fallback;

            if (ImagePixelSize.IsPng(data))
                return "image/png";
            if (IsJpeg(data))
                return "image/jpeg";
            if (ImagePixelSize.IsGif(data))
                return "image/gif";
            if (ImagePixelSize.IsBmp(data))
                return "image/bmp";
            if (ImagePixelSize.IsTiff(data))
                return "image/tiff";
            if (ImagePixelSize.IsEmf(data))
                return "image/x-emf";
            if (ImagePixelSize.IsPlaceableWmf(data) || ImagePixelSize.IsNonPlaceableWmf(data))
                return "image/x-wmf";

            return Fallback;
        }

        /// <summary>
        /// Gets the usual file extension of an image content type.
        /// </summary>
        /// <param name="contentType">The content type.</param>
        /// <returns>The extension, without a dot.</returns>
        internal static string GetExtension(string contentType)
        {
            switch (contentType)
            {
                case "image/png":
                    return "png";
                case "image/gif":
                    return "gif";
                case "image/bmp":
                    return "bmp";
                case "image/tiff":
                    return "tif";
                case "image/x-emf":
                    return "emf";
                case "image/x-wmf":
                    return "wmf";
                default:
                    return "jpg";
            }
        }

        private static bool IsJpeg(byte[] data)
        {
            return data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF;
        }
    }
}
