//TODO: Expand paragraph properties
namespace OpenXMLMailMerge.Core.OpenXMLDocument.Document.Element
{
    /// <summary>
    /// Text element definitions.
    /// </summary>
    public class TextElementConfig
    {
        /// <summary>
        /// Font.
        /// </summary>
        internal string FontName { get; set; } = "Calibri";
        /// <summary>
        /// Font size.
        /// </summary>
        internal int FontSize { get; set; } = 22;
        /// <summary>
        /// Bold option.
        /// </summary>
        internal bool Bold { get; set; } = false;
        /// <summary>
        /// Italic option.
        /// </summary>
        internal bool Italic { get; set; } = false;
        /// <summary>
        /// Strike option.
        /// </summary>
        internal bool Strike { get; set; } = false;
    }
}
