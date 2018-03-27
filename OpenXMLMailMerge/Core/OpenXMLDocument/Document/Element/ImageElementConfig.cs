namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Image element definitions.
    /// </summary>
    public class ImageElementConfig
    {
     
        private ImageElementConfig() { }
        
        /// <summary>
        /// Width
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// Height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// Parse the Image element definition.
        /// </summary>
        /// <param name="arg">String with parameters and values to be parsed.</param>
        public static ImageElementConfig ParseConfig(string arg)
        {
            var config = new Configuration.Configuration().GenerateConfiguration(arg);

            if (config.GenericConfiguration?.Configuration?.Count == 0)
                return null;

            var result = new ImageElementConfig();

            foreach (var item in config.GenericConfiguration?.Configuration)
            {
                foreach (var value in item.Value)
                {
                    if (value.Key == "WIDTH")
                    {
                        int.TryParse(value.Value, out int width);
                        result.Width = width;
                    }

                    if (value.Key == "HEIGHT")
                    {
                        int.TryParse(value.Value, out int height);
                        result.Height = height;
                    }
                }
            }

            return result;
        }
    }
}
