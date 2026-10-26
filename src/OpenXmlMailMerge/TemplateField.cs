#nullable enable

using System.Collections.Generic;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// One MERGEFIELD in a template.
    /// </summary>
    public sealed class TemplateField
    {
        internal TemplateField(string name, IReadOnlyDictionary<string, string> parameters, string partUri, bool isComplex, bool isMergeable)
        {
            Name = name;
            Parameters = parameters;
            PartUri = partUri;
            IsComplex = isComplex;
            IsMergeable = isMergeable;
        }

        /// <summary>
        /// The binding name: the field name without its parameters. A binding with this name, ignoring case, fills the
        /// field.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The parameters written after the name (WIDTH, HEIGHT, STYLE, ...), with keys that ignore case. Empty when the
        /// field has none.
        /// </summary>
        public IReadOnlyDictionary<string, string> Parameters { get; }

        /// <summary>
        /// The part holding the field, for example "/word/header1.xml".
        /// </summary>
        public string PartUri { get; }

        /// <summary>
        /// True when Word stored the field as a complex field (w:fldChar runs) rather than a simple one (w:fldSimple).
        /// </summary>
        public bool IsComplex { get; }

        /// <summary>
        /// False when the field can't be merged even with a binding: it is nested in another field, spread over
        /// paragraphs, or its field code holds content the merge would lose.
        /// </summary>
        public bool IsMergeable { get; }
    }
}
