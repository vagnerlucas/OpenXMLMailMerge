using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Paragraph element definitions.
    /// </summary>
    internal class ParagraphElementConfig
    {
        /// <summary>
        /// Justification option.
        /// </summary>
        internal JustificationValues Justification { get; set; } = JustificationValues.Left;
    }

    /// <summary>
    /// Extension to get the OpenXml justification enum.
    /// </summary>
    internal static class ParagraphElementConfigExtensions
    {
        /// <summary>
        /// Gets the OpenXml Justification enum value
        /// </summary>
        /// <param name="paragraphElementConfig">Paragraph element definition.</param>
        /// <returns>OpenXml Justification.</returns>
        internal static Justification GetValue(this ParagraphElementConfig paragraphElementConfig)
        {
            return new Justification() { Val = paragraphElementConfig.Justification };
        }
    }
}
