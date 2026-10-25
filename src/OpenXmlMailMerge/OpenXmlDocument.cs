using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// The OpenXml document processor.
    /// The document is merged in memory: loading copies the template into a private buffer,
    /// and the template itself (file, stream or byte array) is never modified.
    /// </summary>
    public abstract class OpenXmlDocument
    {
        /// <summary>
        /// Private, expandable buffer holding the package being merged.
        /// </summary>
        private MemoryStream _buffer;

        /// <summary>
        /// Only the document processors of this library derive from this class.
        /// </summary>
        private protected OpenXmlDocument() { }

        /// <summary>
        /// The data dictionary with the merge content and attributes to be processed.
        /// </summary>
        private protected DataDictionary Dictionary { get; } = new DataDictionary();

        /// <summary>
        /// Manages the document content.
        /// </summary>
        private protected DocumentHandler ContentManager { get; set; }

        /// <summary>
        /// Adds a data to the dictionary to be merged.
        /// </summary>
        /// <param name="dataType">The mailmerge data type.</param>
        /// <param name="id">The document's field id.</param>
        /// <param name="data">Data to replace the field.</param>
        public void AddToDictionary(MailMergeDataTypeEnum dataType, string id, object data)
        {
            if (!Dictionary.TryGetValue(dataType, out var value))
            {
                var tmpDictionary = new Dictionary<string, object>();
                value = tmpDictionary;
                Dictionary.Add(dataType, value);
            }

            value.Add(id, data);
        }

        /// <summary>
        /// Opens the package of the concrete document type on the given stream.
        /// </summary>
        /// <param name="stream">Expandable, seekable stream holding the package.</param>
        /// <returns>The content handler of the opened package.</returns>
        private protected abstract DocumentHandler OpenContent(Stream stream);

        /// <summary>
        /// Loads a document from a byte array. The array is copied and never modified.
        /// </summary>
        /// <param name="bytes">Array of bytes with file content.</param>
        /// <returns>OpenXmlDocument processor.</returns>
        public OpenXmlDocument LoadFromBytes(byte[] bytes)
        {
            ArgumentNullException.ThrowIfNull(bytes);

            // Not new MemoryStream(bytes): that stream is fixed-size, fails as soon as the package
            // grows, and writes into the caller's array.
            var buffer = new MemoryStream(bytes.Length);
            buffer.Write(bytes, 0, bytes.Length);

            return Open(buffer);
        }

        /// <summary>
        /// Loads a document from a stream. The stream is read to its end and never modified.
        /// </summary>
        /// <param name="stream">Stream with file content.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>OpenXmlDocument processor.</returns>
        public async Task<OpenXmlDocument> LoadFromStreamAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(stream);

            var buffer = new MemoryStream();
            await stream.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);

            return Open(buffer);
        }

        /// <summary>
        /// Loads a document from file. The file is read once and never modified.
        /// </summary>
        /// <param name="path">Path of the document.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>OpenXmlDocument processor.</returns>
        public async Task<OpenXmlDocument> LoadFromFileAsync(string path, CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);

            await using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            return await LoadFromStreamAsync(file, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Gets the document byte array with every change made so far.
        /// The document stays open and can still be changed afterwards.
        /// </summary>
        /// <returns>Array of byte with document's content.</returns>
        public byte[] GetBytes()
        {
            Flush();
            return _buffer.ToArray();
        }

        /// <summary>
        /// Writes the document, with every change made so far, to a stream.
        /// The document stays open and can still be changed afterwards.
        /// </summary>
        /// <param name="stream">Destination stream. It is written from its current position and not closed.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveToStreamAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(stream);

            Flush();

            // Read the buffer without moving its position, which belongs to the open package.
            var content = _buffer.GetBuffer().AsMemory(0, (int)_buffer.Length);
            await stream.WriteAsync(content, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Writes the document, with every change made so far, to a file. An existing file is overwritten.
        /// </summary>
        /// <param name="path">Path to save.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveToFileAsync(string path, CancellationToken cancellationToken = default)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);

            await using var file = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous);
            await SaveToStreamAsync(file, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Closes the document and releases the in-memory package.
        /// </summary>
        public virtual void Terminate()
        {
            Close();
            Dictionary.Clear();
        }

        /// <summary>
        /// Process the document applying the mail merge.
        /// </summary>
        public virtual void Process()
        {
            Validate();
        }

        /// <summary>
        /// Opens the package held by the buffer, closing any document loaded before.
        /// </summary>
        /// <param name="buffer">Buffer holding the package.</param>
        /// <returns>OpenXmlDocument processor.</returns>
        private OpenXmlDocument Open(MemoryStream buffer)
        {
            Close();

            buffer.Position = 0;
            _buffer = buffer;
            ContentManager = OpenContent(buffer);

            return this;
        }

        /// <summary>
        /// Writes every changed part into the buffer. The package stays open.
        /// </summary>
        private void Flush()
        {
            EnsureLoaded();
            ContentManager.Document.Save();
        }

        /// <summary>
        /// Disposes the package and its buffer.
        /// </summary>
        private void Close()
        {
            ContentManager?.Dispose();
            ContentManager = null;

            _buffer?.Dispose();
            _buffer = null;
        }

        /// <summary>
        /// Ensures a document was loaded.
        /// </summary>
        private void EnsureLoaded()
        {
            if (ContentManager?.Document == null)
                throw new InvalidOperationException("No document loaded. Call LoadFromBytes, LoadFromStreamAsync or LoadFromFileAsync first.");
        }

        /// <summary>
        /// Validates the document, content and dictionary.
        /// </summary>
        private void Validate()
        {
            EnsureLoaded();

            if (Dictionary == null || Dictionary.Count == 0)
                throw new InvalidOperationException("No data to merge");
        }
    }
}
