using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlMailMerge.Fields;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// What the template decides about merged content: the culture of values, the table style, and the text width.
    /// </summary>
    internal static class TemplateFormatting
    {
        /// <summary>
        /// Gets the culture used to format the values merged into a field, from the template's language:
        /// the language set on the field's text, else the document's default language, else the invariant culture.
        /// The machine running the merge never changes the result.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <param name="main">The template's main document part.</param>
        /// <returns>The culture.</returns>
        internal static CultureInfo GetCulture(SimpleField field, MainDocumentPart main)
        {
            var language = field.Descendants<RunProperties>().FirstOrDefault()?.Languages?.Val?.Value
                           ?? main.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.Languages?.Val?.Value;

            if (string.IsNullOrEmpty(language))
                return CultureInfo.InvariantCulture;

            try
            {
                return CultureInfo.GetCultureInfo(language);
            }
            catch (CultureNotFoundException)
            {
                return CultureInfo.InvariantCulture;
            }
        }

        /// <summary>
        /// Finds the table style named by the STYLE parameter of a field, by style id or by style name (ignoring case).
        /// A name with spaces needs a quoted field name: MERGEFIELD "TABLE@STYLE=Grid Table 4".
        /// </summary>
        /// <param name="field">The table field.</param>
        /// <param name="main">The template's main document part.</param>
        /// <returns>The style id, or null when no STYLE is given or the document doesn't define that table style.</returns>
        internal static string FindTableStyleId(SimpleField field, MainDocumentPart main)
        {
            var parameters = FieldParameters.Parse(MergeFieldName.Get(field));
            if (!parameters.TryGetValue("STYLE", out var name) || name.Length == 0)
                return null;

            var style = main.StyleDefinitionsPart?.Styles?.Elements<Style>()
                .FirstOrDefault(s => s.Type?.Value == StyleValues.Table
                                     && (string.Equals(s.StyleId?.Value, name, StringComparison.OrdinalIgnoreCase)
                                         || string.Equals(s.StyleName?.Val?.Value, name, StringComparison.OrdinalIgnoreCase)));

            return style?.StyleId?.Value;
        }

        /// <summary>
        /// Gets the width between the margins of the document's last section, in pixels at 96 DPI (15 twips per pixel).
        /// Headers and footers span the same margins as the body.
        /// The last section is the one in the body's own w:sectPr, which covers the whole document unless it has section breaks.
        /// </summary>
        /// <param name="main">The template's main document part.</param>
        /// <returns>The text width in pixels, or 0 when the page size is not defined.</returns>
        internal static int GetTextWidthInPixels(MainDocumentPart main)
        {
            var section = main.Document.Body?.Elements<SectionProperties>().LastOrDefault();
            var pageWidth = section?.GetFirstChild<PageSize>()?.Width?.Value;
            if (pageWidth == null)
                return 0;

            var margin = section.GetFirstChild<PageMargin>();
            var textWidth = (long)pageWidth.Value - (margin?.Left?.Value ?? 0) - (margin?.Right?.Value ?? 0);

            return textWidth > 0 ? (int)(textWidth / 15) : 0;
        }
    }
}
