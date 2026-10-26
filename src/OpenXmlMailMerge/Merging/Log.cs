using System.Collections.Generic;
using Microsoft.Extensions.Logging;

namespace OpenXmlMailMerge.Merging
{
    /// <summary>
    /// The library's log messages. They carry names and sizes, never merged values: templates often hold personal data.
    /// The source generator skips all formatting when a level is disabled.
    /// </summary>
    internal static partial class Log
    {
        [LoggerMessage(EventId = 1, Level = LogLevel.Debug, Message = "Merging a template with {BindingCount} bindings.")]
        internal static partial void MergeStarting(ILogger logger, int bindingCount);

        [LoggerMessage(EventId = 2, Level = LogLevel.Debug,
            Message = "Merged in {ElapsedMilliseconds} ms: {DocumentBytes} bytes, {UnboundFieldCount} unbound fields, {UnusedBindingCount} unused bindings.")]
        internal static partial void MergeCompleted(ILogger logger, long elapsedMilliseconds, long documentBytes, int unboundFieldCount, int unusedBindingCount);

        [LoggerMessage(EventId = 3, Level = LogLevel.Debug, Message = "Fields without data: {UnboundFields}.")]
        internal static partial void UnboundFields(ILogger logger, IReadOnlyList<string> unboundFields);

        [LoggerMessage(EventId = 4, Level = LogLevel.Debug, Message = "Bindings no field used: {UnusedBindings}.")]
        internal static partial void UnusedBindings(ILogger logger, IReadOnlyList<string> unusedBindings);

        [LoggerMessage(EventId = 5, Level = LogLevel.Warning,
            Message = "The binding {BindingName} matches a field that can't be merged: it is nested in another field, or spread over paragraphs.")]
        internal static partial void BindingNotMerged(ILogger logger, string bindingName);

        [LoggerMessage(EventId = 6, Level = LogLevel.Debug, Message = "Inspected a template: {FieldCount} merge fields.")]
        internal static partial void Inspected(ILogger logger, int fieldCount);
    }
}
