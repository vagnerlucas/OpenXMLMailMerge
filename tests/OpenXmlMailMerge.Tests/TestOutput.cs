using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace OpenXmlMailMerge.Tests
{
    /// <summary>
    /// The folder where tests keep the documents they merge (bin/&lt;configuration&gt;/&lt;tfm&gt;/Output), to be opened
    /// in Word. It is emptied once per test run, before any test class starts.
    /// </summary>
    [TestClass]
    public static class TestOutput
    {
        public static readonly string Directory = Path.Combine(AppContext.BaseDirectory, "Output");

        [AssemblyInitialize]
        public static void Clear(TestContext context)
        {
            if (System.IO.Directory.Exists(Directory))
                System.IO.Directory.Delete(Directory, true);

            System.IO.Directory.CreateDirectory(Directory);
        }

        /// <summary>
        /// Gets the output path for the running test: its name, plus a suffix for data-driven rows, with characters that
        /// aren't valid in file names (on any platform) replaced.
        /// </summary>
        public static string PathFor(TestContext context, string suffix = null)
        {
            var name = suffix == null ? context.TestName : $"{context.TestName}_{suffix}";
            foreach (var c in Path.GetInvalidFileNameChars().Concat(['$', '*', '?', '"', '\\', '/', ':', '<', '>', '|', ' ']))
            {
                name = name?.Replace(c, '_');
            }

            var path = Path.Combine(Directory, name + ".docx");
            context.WriteLine($"Output: {path}");
            return path;
        }
    }
}
