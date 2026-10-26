using System;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXmlMailMerge.Fields
{
    /// <summary>
    /// Reads the field name out of a MERGEFIELD instruction.
    /// </summary>
    internal static class MergeFieldName
    {
        /// <summary>
        /// Gets the field name of a simple field, including its parameters.
        /// </summary>
        /// <param name="field">The field.</param>
        /// <returns>The field name, or null when the field is not a MERGEFIELD.</returns>
        internal static string Get(SimpleField field)
        {
            return Get(field.Instruction?.Value);
        }

        /// <summary>
        /// Gets the field name from a MERGEFIELD instruction, including its parameters.
        /// For example, " MERGEFIELD  IMG@WIDTH=100 \* MERGEFORMAT " gives "IMG@WIDTH=100". A name with spaces is quoted.
        /// </summary>
        /// <param name="instruction">The field instruction.</param>
        /// <returns>The field name, or null when the instruction is not a MERGEFIELD.</returns>
        internal static string Get(string instruction)
        {
            const string keyword = "MERGEFIELD";

            instruction = instruction?.Trim();
            if (instruction == null
                || instruction.Length <= keyword.Length
                || !instruction.StartsWith(keyword, StringComparison.OrdinalIgnoreCase)
                || !char.IsWhiteSpace(instruction[keyword.Length]))
                return null;

            var rest = instruction[keyword.Length..].TrimStart();
            string name;
            if (rest.StartsWith('"'))
            {
                var end = rest.IndexOf('"', 1);
                name = end < 0 ? rest[1..] : rest[1..end];
            }
            else
            {
                var end = rest.IndexOfAny([' ', '\t']);
                name = end < 0 ? rest : rest[..end];
            }

            return name.Length == 0 ? null : name;
        }

        /// <summary>
        /// Gets the binding name of a MERGEFIELD instruction: its field name without parameters.
        /// </summary>
        /// <param name="instruction">The field instruction.</param>
        /// <returns>The binding name, or null when the instruction is not a MERGEFIELD.</returns>
        internal static string GetBindingName(string instruction)
        {
            var name = Get(instruction);
            return name == null ? null : FieldParameters.GetBindingName(name);
        }
    }
}
