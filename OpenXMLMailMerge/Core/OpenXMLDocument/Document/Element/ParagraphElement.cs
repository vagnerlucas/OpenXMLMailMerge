using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Paragraph element class.
    /// </summary>
    internal class ParagraphElement
    {
        /// <summary>
        /// Creates the OpenXml ParagraphProperties.
        /// </summary>
        /// <param name="paragraphElementConfig">Paragraph definitions.</param>
        /// <returns>Paragraph OpenXmlElement.</returns>
        internal static OpenXmlElement CreateParagraphProperties(ParagraphElementConfig paragraphElementConfig = null)
        {
            if (paragraphElementConfig == null)
                paragraphElementConfig = new ParagraphElementConfig();

            return new ParagraphProperties() { Justification = paragraphElementConfig.GetValue() };
        }

        /// <summary>
        /// Creates a Paragraph OpenXmlElement.
        /// </summary>
        /// <param name="text">Text value.</param>
        /// <param name="justificationValue">Justification option.</param>
        /// <param name="textElementConfig">Text properties definitions.</param>
        /// <returns></returns>
        internal static OpenXmlElement CreateParagraph(string text = null, JustificationValues justificationValue = JustificationValues.Left, TextElementConfig textElementConfig = null)
        {
            var paragraphConfig = new ParagraphElementConfig() { Justification = justificationValue };
            var paragraphProperties = CreateParagraphProperties(paragraphConfig);
            var result = new Paragraph() { ParagraphProperties = paragraphProperties as ParagraphProperties};

            if (text != null)
            {
                if (textElementConfig == null)
                    textElementConfig = new TextElementConfig();

                var runProperties = RunElement.CreateRunProperties(textElementConfig);
                var textElement = RunElement.CreateRunElement(text, textElementConfig);
                result.Append(textElement);
            }

            return result;
        }
    }
}
