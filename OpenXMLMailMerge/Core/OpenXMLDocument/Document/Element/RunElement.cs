using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Run element class.
    /// </summary>
    internal class RunElement
    {
        /// <summary>
        /// Creates Run element properties.
        /// </summary>
        /// <param name="textElementConfig">Text properties definitions.</param>
        /// <returns>RunProperties OpenXmlElement.</returns>
        internal static OpenXmlElement CreateRunProperties(TextElementConfig textElementConfig = null) 
        {
            if (textElementConfig == null)
                textElementConfig = new TextElementConfig();

            return new RunProperties(new RunFonts() { Ascii = textElementConfig.FontName }, new FontSize() { Val = textElementConfig.FontSize.ToString() })
            {
                Bold = textElementConfig.Bold ? new Bold() : null,
                Italic = textElementConfig.Italic ? new Italic() : null,
                Strike = textElementConfig.Strike ? new Strike() : null
            };
        }

        /// <summary>
        /// Creates Run OpenXmlElement wth a text (optionally).
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="textElementConfig">Text properties definitions.</param>
        /// <returns>Run OpenXmlElement.</returns>
        internal static OpenXmlElement CreateRunElement(string text = null, TextElementConfig textElementConfig = null)
        {
            var result = new Run();

            if (textElementConfig != null)
            {
                var runProperties = CreateRunProperties(textElementConfig);
                result.RunProperties = runProperties as RunProperties;
            }

            if (text == null)
                return result;

            var textElement = TextElement.CreateText(text);
            result.Append(textElement);
            return result;
        }
    }
}
