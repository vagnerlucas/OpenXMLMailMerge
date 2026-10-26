#nullable enable

using System;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// A merge failed because of the template, not because of an invalid argument: for example, the template isn't a
    /// valid .docx document.
    /// </summary>
    public sealed class MergeException : Exception
    {
        /// <summary>
        /// Creates the exception.
        /// </summary>
        /// <param name="message">What went wrong.</param>
        public MergeException(string message) : base(message)
        {
        }

        /// <summary>
        /// Creates the exception.
        /// </summary>
        /// <param name="message">What went wrong.</param>
        /// <param name="innerException">The underlying error.</param>
        public MergeException(string message, Exception? innerException) : base(message, innerException)
        {
        }

        /// <summary>
        /// Creates the exception for a specific field.
        /// </summary>
        /// <param name="message">What went wrong.</param>
        /// <param name="fieldName">The field involved.</param>
        /// <param name="partUri">The part holding the field, for example "/word/header1.xml".</param>
        /// <param name="innerException">The underlying error, if any.</param>
        public MergeException(string message, string? fieldName, string? partUri, Exception? innerException = null)
            : base(message, innerException)
        {
            FieldName = fieldName;
            PartUri = partUri;
        }

        /// <summary>
        /// The field involved, when the error concerns one field.
        /// </summary>
        public string? FieldName { get; }

        /// <summary>
        /// The part holding the field, for example "/word/header1.xml", when the error concerns one field.
        /// </summary>
        public string? PartUri { get; }
    }
}
