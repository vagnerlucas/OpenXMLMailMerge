#nullable enable

namespace OpenXmlMailMerge
{
    /// <summary>
    /// What a merge does with a field that gets no data.
    /// </summary>
    public enum UnboundFieldBehavior
    {
        /// <summary>
        /// Leave the field in the document, showing its placeholder (for example «Name»). The default.
        /// </summary>
        Keep,
        /// <summary>
        /// Remove the field and its placeholder. A field nested in another field, or spread over paragraphs, can't be
        /// removed safely and stays.
        /// </summary>
        Remove,
        /// <summary>
        /// Fail the merge with a <see cref="MergeException"/> naming the first such field and its part. Nothing is
        /// written to the output.
        /// </summary>
        Throw
    }
}
