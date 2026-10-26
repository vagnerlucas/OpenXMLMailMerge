using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Fields
{
    /// <summary>
    /// Lists the instructions of every field under an element: simple fields (w:fldSimple) and complex fields
    /// (w:fldChar runs), nested and multi-paragraph ones included.
    /// </summary>
    internal static class FieldInstructions
    {
        /// <summary>
        /// Gets every field instruction under an element, in document order. A complex field is listed when its
        /// instruction ends (at its separate or end marker), so a field nested in another one's instruction comes first.
        /// </summary>
        /// <param name="root">Element to search, such as a document, header or footer.</param>
        /// <returns>
        /// Each instruction, for example " MERGEFIELD  NAME \* MERGEFORMAT ", with its simple field; the field is null for
        /// a complex field.
        /// </returns>
        internal static IEnumerable<(string Instruction, SimpleField Field)> Find(OpenXmlElement root)
        {
            // One entry per complex field still open; null once its instruction is complete.
            var open = new Stack<StringBuilder>();

            foreach (var element in root.Descendants())
            {
                switch (element)
                {
                    case SimpleField field:
                        yield return (field.Instruction?.Value ?? string.Empty, field);
                        break;

                    case FieldCode code when open.Count > 0 && open.Peek() != null:
                        open.Peek().Append(code.Text);
                        break;

                    case FieldChar marker when marker.FieldCharType?.Value == FieldCharValues.Begin:
                        open.Push(new StringBuilder());
                        break;

                    case FieldChar marker when marker.FieldCharType?.Value == FieldCharValues.Separate && open.Count > 0:
                        var instruction = open.Pop();
                        open.Push(null);
                        if (instruction != null)
                            yield return (instruction.ToString(), null);
                        break;

                    case FieldChar marker when marker.FieldCharType?.Value == FieldCharValues.End && open.Count > 0:
                        var unfinished = open.Pop();
                        if (unfinished != null)
                            yield return (unfinished.ToString(), null);
                        break;
                }
            }
        }
    }
}
