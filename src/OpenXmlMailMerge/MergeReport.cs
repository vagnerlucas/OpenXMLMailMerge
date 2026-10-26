#nullable enable

using System.Collections.Generic;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// What a merge matched and what it didn't, to catch a typo in the template or in the code that feeds it.
    /// This is information, not an error.
    /// </summary>
    public sealed class MergeReport
    {
        internal MergeReport(IReadOnlyList<string> unboundFields, IReadOnlyList<string> unusedBindings)
        {
            UnboundFields = unboundFields;
            UnusedBindings = unusedBindings;
        }

        /// <summary>
        /// Fields that got no data, by binding name, without duplicates: headers first, then footers, then the body.
        /// Includes bound fields that can't be merged (nested in another field, or spread over paragraphs). Listed even
        /// when <see cref="MergeOptions.UnboundFields"/> is Remove and they were taken out of the document.
        /// </summary>
        public IReadOnlyList<string> UnboundFields { get; }

        /// <summary>
        /// Binding names that no field in the template used, in the order they were added.
        /// </summary>
        public IReadOnlyList<string> UnusedBindings { get; }
    }
}
