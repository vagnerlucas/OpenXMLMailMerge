using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using Microsoft.Extensions.Logging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using W = DocumentFormat.OpenXml.Wordprocessing;
using static OpenXmlMailMerge.Tests.TestDocuments;

namespace OpenXmlMailMerge.Tests
{
    [TestClass]
    public class LoggingTests
    {
        /// <summary>
        /// A logger that keeps every entry in memory, at every level.
        /// </summary>
        private sealed class CapturingLogger<T> : ILogger<T>
        {
            public List<(LogLevel Level, int EventId, string Message)> Entries { get; } = [];

            public IDisposable BeginScope<TState>(TState state) where TState : notnull
            {
                return null;
            }

            public bool IsEnabled(LogLevel logLevel)
            {
                return true;
            }

            public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter)
            {
                Entries.Add((logLevel, eventId.Id, formatter(state, exception)));
            }
        }

        private static (DocumentMerger Merger, CapturingLogger<DocumentMerger> Logger) Create()
        {
            var logger = new CapturingLogger<DocumentMerger>();
            return (new DocumentMerger(logger: logger), logger);
        }

        [TestMethod]
        public void Merge_LogsStartAndCompletion_AtDebug()
        {
            var (merger, logger) = Create();

            var result = merger.Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis"));

            var start = logger.Entries.Single(entry => entry.EventId == 1);
            var completion = logger.Entries.Single(entry => entry.EventId == 2);
            Assert.AreEqual(LogLevel.Debug, start.Level);
            Assert.AreEqual(LogLevel.Debug, completion.Level);
            StringAssert.Contains(start.Message, "1 bindings");
            StringAssert.Contains(completion.Message, $"{result.Document.Length} bytes, 4 unbound fields, 0 unused bindings");
        }

        [TestMethod]
        public async Task MergeAsync_LogsCompletion()
        {
            var (merger, logger) = Create();
            using var output = new MemoryStream();

            await merger.MergeAsync(new MemoryStream(File.ReadAllBytes(BindFile)), output, new MergeData().Text("BIND", "Reis"));

            StringAssert.Contains(logger.Entries.Single(entry => entry.EventId == 2).Message, $"{output.Length} bytes");
        }

        [TestMethod]
        public void Merge_LogsUnboundFieldsAndUnusedBindings()
        {
            var (merger, logger) = Create();

            merger.Merge(File.ReadAllBytes(BindFile), new MergeData().Text("BIND", "Reis").Text("Adress", "Main Street"));

            StringAssert.Contains(logger.Entries.Single(entry => entry.EventId == 3).Message, "HEADER, IMG, FOOTER, TABLE");
            StringAssert.Contains(logger.Entries.Single(entry => entry.EventId == 4).Message, "Adress");
            Assert.IsFalse(logger.Entries.Any(entry => entry.Level >= LogLevel.Warning), "Unbound fields and unused bindings are not warnings.");
        }

        [TestMethod]
        public void Merge_BindingThatCantBeMerged_LogsWarning()
        {
            // { IF { MERGEFIELD NAME } = "x" "yes" "no" }: NAME is bound, but nested, so it isn't merged.
            OpenXmlElement[] field =
            [
                FieldMarker(W.FieldCharValues.Begin),
                new W.Run(new W.FieldCode(" IF ") { Space = SpaceProcessingModeValues.Preserve }),
                .. ComplexField("NAME"),
                new W.Run(new W.FieldCode(" = \"x\" \"yes\" \"no\" ") { Space = SpaceProcessingModeValues.Preserve }),
                FieldMarker(W.FieldCharValues.Separate),
                new W.Run(new W.Text("no")),
                FieldMarker(W.FieldCharValues.End)
            ];
            var (merger, logger) = Create();

            merger.Merge(CreateDocument([new W.Paragraph(field)]), new MergeData().Text("NAME", "Jane"));

            var warning = logger.Entries.Single(entry => entry.Level == LogLevel.Warning);
            Assert.AreEqual(5, warning.EventId);
            StringAssert.Contains(warning.Message, "NAME");
        }

        [TestMethod]
        public void Logs_NeverContainMergedValues()
        {
            var (merger, logger) = Create();
            var data = new MergeData()
                .Text("BIND", "Secret value 42")
                .Table("TABLE", new TableValue(["Salary"], [["Confidential 9000"]]));

            merger.Merge(File.ReadAllBytes(BindFile), data);
            merger.Inspect(File.ReadAllBytes(BindFile));

            Assert.IsTrue(logger.Entries.Count > 0);
            foreach (var entry in logger.Entries)
            {
                Assert.IsFalse(entry.Message.Contains("Secret") || entry.Message.Contains("Confidential"), entry.Message);
            }
        }

        [TestMethod]
        public void Inspect_LogsFieldCount()
        {
            var (merger, logger) = Create();

            merger.Inspect(File.ReadAllBytes(BindFile));

            StringAssert.Contains(logger.Entries.Single(entry => entry.EventId == 6).Message, "7 merge fields");
        }
    }
}
