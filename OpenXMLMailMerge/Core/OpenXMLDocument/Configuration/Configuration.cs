using System.Collections.Generic;
using System.Linq;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Configuration
{
    /// <summary>
    /// User defined processor configuration to handle document's field with parameters
    /// Example of parameter (without spaces): field @ parameter1=0 @ parameter2=1...
    /// </summary>
    internal sealed class Configuration
    {
        /// <summary>
        /// Generic (dynamic) configuration. Free to be defined.
        /// </summary>
        internal GenericConfiguration GenericConfiguration { get; private set; } = new GenericConfiguration();

        /// <summary>
        /// Generates configuration based on parameters set in field tag.
        /// </summary>
        /// <param name="arg">String to be parsed.</param>
        /// <returns>A dictionary with parameters and values defined by user.</returns>
        internal Configuration GenerateConfiguration(string arg)
        {
            arg = arg
                    .Replace(StringConst.SPECIAL_STR_L, string.Empty)
                    .Replace(StringConst.SPECIAL_STR_R, string.Empty);

            GenericConfiguration.Configuration = new Dictionary<string, object>();
            foreach (var item in GetParameters(arg))
            {
                GenericConfiguration.Configuration.Add(item.Key.ToUpper().Trim(), item.Value);
            }

            return this;
        }

        /// <summary>
        /// Gets the parameters of the text tag by splitting with defined delimiters
        /// A parameter should have the format: field@parameter=value.
        /// </summary>
        /// <param name="arg">String to be parsed.</param>
        /// <returns>A dictionary with parameters and values.</returns>
        private Dictionary<string, Dictionary<string, object>> GetParameters(string arg)
        {
            var result = new Dictionary<string, Dictionary<string, object>>();
            string masterkey = string.Empty;

            var splitValue = arg.Split(StringConst.ID_DELIMITER_CHAR);

            if (splitValue.Length > 1)
            {
                foreach (var item in splitValue)
                {
                    if (item.Contains(StringConst.KEY_VALUE_DELIMITER))
                    {
                        var toSplit = item.Split(StringConst.KEY_VALUE_DELIMITER);
                        if (toSplit.Length <= 0) continue;
                        var key = toSplit[0].ToUpper();
                        string value = toSplit[1];
                        result[masterkey].Add(key, value);
                    }
                    else
                    {
                        masterkey = item.Trim();
                        result.Add(masterkey, new Dictionary<string, object>());
                    }
                }
            }

            return result;
        }
    }
}
