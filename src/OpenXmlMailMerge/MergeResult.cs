#nullable enable

namespace OpenXmlMailMerge
{
    /// <summary>
    /// The result of an in-memory merge: the merged document and the report.
    /// </summary>
    public sealed class MergeResult
    {
        internal MergeResult(byte[] document, MergeReport report)
        {
            Document = document;
            Report = report;
        }

        /// <summary>
        /// The merged .docx document.
        /// </summary>
        public byte[] Document { get; }

        /// <summary>
        /// What the merge matched and what it didn't.
        /// </summary>
        public MergeReport Report { get; }
    }
}
