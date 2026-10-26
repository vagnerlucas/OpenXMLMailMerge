using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// Puts merged content where a field was, and removes the field.
    /// </summary>
    internal static class FieldReplacement
    {
        /// <summary>
        /// Replaces a field with a run holding inline content (text or an image), at the field's position.
        /// The run keeps the formatting of the field's first result run.
        /// </summary>
        /// <param name="field">The field to replace.</param>
        /// <param name="content">The inline content.</param>
        internal static void ReplaceWithInline(SimpleField field, OpenXmlElement content)
        {
            var run = new Run();

            var properties = field.Descendants<RunProperties>().FirstOrDefault();
            if (properties != null)
                run.Append(properties.CloneNode(true));

            run.Append(content);

            field.InsertBeforeSelf(run);
            field.Remove();
        }

        /// <summary>
        /// Puts a block element (such as a table) where a field is, and removes the field.
        /// A block can't be inside a paragraph, so:
        /// <list type="bullet">
        /// <item>when the field is the only content of its paragraph, the block replaces the paragraph;</item>
        /// <item>otherwise the block goes right after the paragraph, which keeps its other content.</item>
        /// </list>
        /// The paragraph is kept when it carries a section break, or when it is the last element of its container
        /// (table cells, headers and footers must end with a paragraph). For the same reason, an empty paragraph is
        /// added when the block would otherwise end its container.
        /// </summary>
        /// <param name="field">The field to replace.</param>
        /// <param name="block">The block element.</param>
        internal static void ReplaceWithBlock(SimpleField field, OpenXmlElement block)
        {
            var paragraph = field.Ancestors<Paragraph>().FirstOrDefault();
            if (paragraph == null)
                return;

            field.Remove();

            var isEmpty = paragraph.Descendants<Text>().All(t => string.IsNullOrWhiteSpace(t.Text))
                          && !paragraph.Descendants<Drawing>().Any()
                          && !paragraph.Descendants<SimpleField>().Any();
            if (!isEmpty)
            {
                paragraph.InsertAfterSelf(block);
            }
            else
            {
                paragraph.InsertBeforeSelf(block);
                if (paragraph.NextSibling() != null && !paragraph.Descendants<SectionProperties>().Any())
                    paragraph.Remove();
            }

            if (block.NextSibling() == null)
                block.InsertAfterSelf(new Paragraph());
        }
    }
}
