#nullable enable

using System.Collections.Generic;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// The merge fields of a template, as found by <see cref="DocumentMerger.Inspect"/>.
    /// </summary>
    public sealed class TemplateInfo
    {
        internal TemplateInfo(IReadOnlyList<TemplateField> fields)
        {
            Fields = fields;
        }

        /// <summary>
        /// Every MERGEFIELD, one entry per occurrence: headers first, then footers, then the body, each in document order.
        /// </summary>
        public IReadOnlyList<TemplateField> Fields { get; }
    }
}
