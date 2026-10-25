using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Fields
{
    /// <summary>
    /// Rewrites complex fields as simple fields.
    /// Word stores a field either as a simple field (one w:fldSimple element) or as a complex field spread over sibling runs:
    /// <code>
    /// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="begin"/&gt;&lt;/w:r&gt;
    /// &lt;w:r&gt;&lt;w:instrText&gt; MERGEFIELD NAME &lt;/w:instrText&gt;&lt;/w:r&gt;   (one or more runs)
    /// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="separate"/&gt;&lt;/w:r&gt;              (optional)
    /// &lt;w:r&gt;&lt;w:t&gt;«NAME»&lt;/w:t&gt;&lt;/w:r&gt;                              (the field result, zero or more runs)
    /// &lt;w:r&gt;&lt;w:fldChar w:fldCharType="end"/&gt;&lt;/w:r&gt;
    /// </code>
    /// Both forms mean the same to Word. A converted field keeps its instruction and its result runs.
    /// </summary>
    internal static class ComplexFields
    {
        /// <summary>
        /// Converts the complex fields under an element into simple fields.
        /// A field is converted only when it is safe to do so: all its runs are siblings (it doesn't span paragraphs),
        /// it holds no nested field and isn't nested in another, its instruction is made of plain instruction runs,
        /// and <paramref name="shouldConvert"/> accepts its instruction. Any other field is left as it is.
        /// </summary>
        /// <param name="root">Element to search, such as a document, header or footer.</param>
        /// <param name="shouldConvert">Decides, from the field instruction, whether the field is converted.</param>
        internal static void ConvertToSimpleFields(OpenXmlElement root, Func<string, bool> shouldConvert)
        {
            // Every element holding the run that starts a field: paragraphs, hyperlinks, and so on.
            var parents = root.Descendants<FieldChar>()
                .Where(c => c.FieldCharType?.Value == FieldCharValues.Begin)
                .Select(c => c.Parent?.Parent)
                .Where(p => p != null)
                .Distinct()
                .ToList();

            foreach (var parent in parents)
                ConvertInParent(parent, shouldConvert);
        }

        private static void ConvertInParent(OpenXmlElement parent, Func<string, bool> shouldConvert)
        {
            var children = parent.ChildElements.ToList();

            // Pair each top-level begin with its end, tracking depth to see nested fields.
            var fields = new List<(int Begin, int End)>();
            var depth = 0;
            var begin = -1;
            var nested = false;
            for (var i = 0; i < children.Count; i++)
            {
                var type = GetMarkerType(children[i]);
                if (type == FieldCharValues.Begin)
                {
                    if (depth == 0)
                    {
                        begin = i;
                        nested = false;
                    }
                    else
                    {
                        nested = true;
                    }

                    depth++;
                }
                else if (type == FieldCharValues.End && depth > 0)
                {
                    depth--;
                    if (depth == 0 && !nested)
                        fields.Add((begin, i));
                }
            }

            // A begin without its end in this parent spans paragraphs: it was never added, so it stays as it is.
            foreach (var (fieldBegin, fieldEnd) in fields)
                TryConvert(children, fieldBegin, fieldEnd, shouldConvert);
        }

        private static void TryConvert(List<OpenXmlElement> children, int begin, int end, Func<string, bool> shouldConvert)
        {
            var separate = -1;
            for (var i = begin + 1; i < end; i++)
            {
                if (GetMarkerType(children[i]) != FieldCharValues.Separate)
                    continue;

                separate = i;
                break;
            }

            var instructionEnd = separate >= 0 ? separate : end;
            var instruction = new StringBuilder();
            for (var i = begin + 1; i < instructionEnd; i++)
            {
                switch (children[i])
                {
                    case ProofError:
                        continue;
                    case Run run when run.ChildElements.All(c => c is RunProperties or FieldCode):
                        foreach (var code in run.Elements<FieldCode>())
                            instruction.Append(code.Text);
                        continue;
                    default:
                        // Anything else (bookmarks, deleted instruction text, ...) would be lost: leave the field alone.
                        return;
                }
            }

            if (!shouldConvert(instruction.ToString()))
                return;

            var field = new SimpleField { Instruction = instruction.ToString() };
            if (separate >= 0)
            {
                for (var i = separate + 1; i < end; i++)
                {
                    children[i].Remove();
                    field.Append(children[i]);
                }
            }

            // Remove the markers and the instruction; the result runs now belong to the new field.
            children[begin].InsertBeforeSelf(field);
            for (var i = begin; i <= instructionEnd; i++)
                children[i].Remove();

            // Without a separate marker, instructionEnd is the end marker and is already removed.
            if (separate >= 0)
                children[end].Remove();
        }

        /// <summary>
        /// Gets the type of a field marker run: a run that holds a w:fldChar and nothing but its properties.
        /// </summary>
        /// <returns>The marker type, or null when the element is not a marker run.</returns>
        private static FieldCharValues? GetMarkerType(OpenXmlElement element)
        {
            if (element is not Run run || !run.ChildElements.All(c => c is RunProperties or FieldChar))
                return null;

            return run.GetFirstChild<FieldChar>()?.FieldCharType?.Value;
        }
    }
}
