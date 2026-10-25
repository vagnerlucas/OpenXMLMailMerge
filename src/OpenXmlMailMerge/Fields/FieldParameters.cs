using System;
using System.Collections.Generic;

namespace OpenXmlMailMerge.Fields
{
    /// <summary>
    /// Parses the parameters written in a field name: <c>name@KEY=value@KEY=value</c>,
    /// for example <c>IMG@WIDTH=120@HEIGHT=80</c>.
    /// </summary>
    internal static class FieldParameters
    {
        /// <summary>
        /// Gets the binding name of a field: its name without the parameters.
        /// </summary>
        /// <param name="fieldName">The field name, for example "IMG@WIDTH=120".</param>
        /// <returns>The binding name, for example "IMG".</returns>
        internal static string GetBindingName(string fieldName)
        {
            var delimiter = fieldName.IndexOf(StringConst.IdDelimiterChar);
            return delimiter < 0 ? fieldName : fieldName[..delimiter];
        }

        /// <summary>
        /// Parses the parameters of a field name. Keys are compared ordinally, ignoring case, so "width" and "WIDTH" are
        /// the same key in every culture. A value is everything after the first '='. Segments without a key or without
        /// '=' are ignored, and a repeated key keeps its last value.
        /// </summary>
        /// <param name="fieldName">The field name, for example "IMG@WIDTH=120@HEIGHT=80".</param>
        /// <returns>The parameters by key; empty when the field has none.</returns>
        internal static IReadOnlyDictionary<string, string> Parse(string fieldName)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(fieldName))
                return result;

            // The first segment is the binding name.
            foreach (var segment in fieldName.Split(StringConst.IdDelimiterChar)[1..])
            {
                var assignment = segment.IndexOf(StringConst.KeyValueDelimiter);
                if (assignment <= 0)
                    continue;

                var key = segment[..assignment].Trim();
                if (key.Length == 0)
                    continue;

                result[key] = segment[(assignment + 1)..].Trim();
            }

            return result;
        }
    }
}
