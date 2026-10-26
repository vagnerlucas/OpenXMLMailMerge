#nullable enable

using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OpenXmlMailMerge.Merging;

namespace OpenXmlMailMerge
{
    /// <summary>
    /// Merges data into Word (.docx) templates. An instance holds no state between merges: it is safe to reuse and to
    /// share between threads, for example as a singleton in dependency injection.
    /// </summary>
    public sealed class DocumentMerger
    {
        private readonly MergeOptions _options;
        private readonly ILogger _logger;

        /// <summary>
        /// Creates a merger.
        /// </summary>
        /// <param name="options">The merge options; null for the defaults.</param>
        /// <param name="logger">
        /// Where to log, through any Microsoft.Extensions.Logging provider; null to log nothing. Merges log at Debug level,
        /// plus a Warning when a binding matches a field that can't be merged. Logs carry field names and sizes, never
        /// merged values.
        /// </param>
        /// <exception cref="ArgumentOutOfRangeException">An option is out of range.</exception>
        public DocumentMerger(MergeOptions? options = null, ILogger<DocumentMerger>? logger = null)
        {
            _logger = logger ?? (ILogger)NullLogger.Instance;

            options ??= new MergeOptions();

            if (!Enum.IsDefined(options.UnboundFields))
                throw new ArgumentOutOfRangeException(nameof(options), options.UnboundFields, "Unknown UnboundFields behavior.");

            if (options.MaxPartCharacters < 0)
                throw new ArgumentOutOfRangeException(nameof(options), options.MaxPartCharacters, "MaxPartCharacters can't be negative.");

            _options = options;
        }

        /// <summary>
        /// Merges data into a template held in memory. The template is copied, never modified.
        /// </summary>
        /// <param name="template">The .docx template.</param>
        /// <param name="data">The data to merge.</param>
        /// <returns>The merged document and the report.</returns>
        /// <exception cref="ArgumentNullException">The data is null.</exception>
        /// <exception cref="MergeException">
        /// The template isn't a valid .docx document, a part is larger than <see cref="MergeOptions.MaxPartCharacters"/>,
        /// fields got no data and <see cref="MergeOptions.UnboundFields"/> is Throw, or the result isn't valid and
        /// <see cref="MergeOptions.ValidateOutput"/> is set.
        /// </exception>
        public MergeResult Merge(ReadOnlySpan<byte> template, MergeData data)
        {
            ArgumentNullException.ThrowIfNull(data);

            var started = Stopwatch.GetTimestamp();
            Log.MergeStarting(_logger, data.Names.Count);

            MergeSession session;
            try
            {
                session = MergeSession.Open(template, _options);
            }
            catch (Exception exception) when (exception is not OperationCanceledException and not MergeException)
            {
                throw InvalidTemplate(exception);
            }

            using (session)
            {
                session.Merge(data);
                var report = session.Complete();
                LogCompleted(started, session.Length, report, data);
                return new MergeResult(session.GetBytes(), report);
            }
        }

        /// <summary>
        /// Merges data into a template read from a stream, and writes the merged document to another stream.
        /// The template stream is read to its end; neither stream is closed.
        /// </summary>
        /// <param name="template">The .docx template.</param>
        /// <param name="output">Where the merged document is written, from its current position.</param>
        /// <param name="data">The data to merge.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>The report.</returns>
        /// <exception cref="ArgumentNullException">A stream or the data is null.</exception>
        /// <exception cref="MergeException">
        /// The template isn't a valid .docx document, a part is larger than <see cref="MergeOptions.MaxPartCharacters"/>,
        /// fields got no data and <see cref="MergeOptions.UnboundFields"/> is Throw, or the result isn't valid and
        /// <see cref="MergeOptions.ValidateOutput"/> is set.
        /// </exception>
        public async Task<MergeReport> MergeAsync(Stream template, Stream output, MergeData data, CancellationToken cancellationToken = default)
        {
            ArgumentNullException.ThrowIfNull(template);
            ArgumentNullException.ThrowIfNull(output);
            ArgumentNullException.ThrowIfNull(data);

            var started = Stopwatch.GetTimestamp();
            Log.MergeStarting(_logger, data.Names.Count);

            MergeSession session;
            try
            {
                session = await MergeSession.OpenAsync(template, _options, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception exception) when (exception is not OperationCanceledException and not MergeException)
            {
                throw InvalidTemplate(exception);
            }

            using (session)
            {
                session.Merge(data);
                var report = session.Complete();
                await session.WriteToAsync(output, cancellationToken).ConfigureAwait(false);
                LogCompleted(started, session.Length, report, data);
                return report;
            }
        }

        /// <summary>
        /// Lists the merge fields of a template held in memory, without merging anything: their names, parameters, parts,
        /// and whether they can be merged. Useful to check a template before merging, or to show its fields in a UI.
        /// </summary>
        /// <param name="template">The .docx template.</param>
        /// <returns>The template's fields.</returns>
        /// <exception cref="MergeException">
        /// The template isn't a valid .docx document, or a part is larger than <see cref="MergeOptions.MaxPartCharacters"/>.
        /// </exception>
        public TemplateInfo Inspect(ReadOnlySpan<byte> template)
        {
            MergeSession session;
            try
            {
                session = MergeSession.Open(template, _options);
            }
            catch (Exception exception) when (exception is not OperationCanceledException and not MergeException)
            {
                throw InvalidTemplate(exception);
            }

            using (session)
            {
                var info = session.Inspect();
                Log.Inspected(_logger, info.Fields.Count);
                return info;
            }
        }

        private void LogCompleted(long started, long documentBytes, MergeReport report, MergeData data)
        {
            Log.MergeCompleted(_logger, (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds, documentBytes,
                report.UnboundFields.Count, report.UnusedBindings.Count);

            if (report.UnboundFields.Count > 0)
                Log.UnboundFields(_logger, report.UnboundFields);

            if (report.UnusedBindings.Count > 0)
                Log.UnusedBindings(_logger, report.UnusedBindings);

            // A field that got no data although a binding has its name is one the engine can't merge: easy to miss.
            foreach (var name in report.UnboundFields)
            {
                if (data.TryGetBinding(name, out _, out _))
                    Log.BindingNotMerged(_logger, name);
            }
        }

        private static MergeException InvalidTemplate(Exception exception)
        {
            return new MergeException("The template isn't a valid .docx document.", exception);
        }
    }
}
