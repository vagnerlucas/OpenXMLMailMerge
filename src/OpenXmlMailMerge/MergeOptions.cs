#nullable enable

using System.Globalization;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// Options for <see cref="DocumentMerger"/>. Immutable: create a modified copy with a <c>with</c> expression.
    /// </summary>
    public sealed record MergeOptions
    {
        /// <summary>
        /// The default for <see cref="MaxPartCharacters"/>: 50 million characters, far beyond any real document part.
        /// </summary>
        public const long DefaultMaxPartCharacters = 50_000_000;

        /// <summary>
        /// What to do with a field that gets no data. Default: <see cref="UnboundFieldBehavior.Keep"/>.
        /// </summary>
        public UnboundFieldBehavior UnboundFields { get; init; } = UnboundFieldBehavior.Keep;

        /// <summary>
        /// The culture numbers and dates are formatted with. Null (the default) uses the template's language: the
        /// language set on the field's text, else the document's default language, else the invariant culture.
        /// </summary>
        public CultureInfo? Culture { get; init; }

        /// <summary>
        /// The largest size, in characters, that any XML part of the template may have once decompressed. A larger part
        /// fails the merge with a <see cref="MergeException"/>. This stops decompression bombs: small templates whose
        /// XML inflates to gigabytes. 0 disables the check. Default: <see cref="DefaultMaxPartCharacters"/>.
        /// </summary>
        public long MaxPartCharacters { get; init; } = DefaultMaxPartCharacters;

        /// <summary>
        /// Validates the merged document against the Open XML schema, and fails the merge with a
        /// <see cref="MergeException"/> instead of returning an invalid document. Nothing is written to the output then.
        /// Validation reads the whole document, so it makes merges noticeably slower. Default: false.
        /// </summary>
        public bool ValidateOutput { get; init; }
    }
}
