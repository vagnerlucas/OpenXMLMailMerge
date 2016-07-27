using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Image element definitions.
    /// </summary>
    public class ImageElementConfig
    {
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
        public void ParseConfig(string arg)
        {
            var config = new Configuration.Configuration().GenerateConfiguration(arg);
            var imageElementConfig = new ImageElementConfig();
            foreach (var item in config.GenericConfiguration.Configuration)
            {
                foreach (var value in item.Value)
                {
                    if (value.Key == "WIDTH")
                    {
                        var width = 0;
                        Int32.TryParse(value.Value, out width);
                        Width = width;
                    }

                    if (value.Key == "HEIGHT")
                    {
                        var height = 0;
                        Int32.TryParse(value.Value, out height);
                        Height = height;
                    }
                }
            }
        }
    }
}
