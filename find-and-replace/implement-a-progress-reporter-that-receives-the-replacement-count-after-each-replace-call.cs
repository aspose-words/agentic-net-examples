using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing; // Required by Aspose.Words for font/color types
using Newtonsoft.Json; // Included as a required package

namespace FindAndReplaceProgressDemo
{
    // Simple progress reporter that receives the replacement count after each Replace call.
    public class ReplacementProgressReporter
    {
        public void Report(int replacementCount)
        {
            Console.WriteLine($"Replacements performed in this step: {replacementCount}");
        }
    }

    // Callback that logs each match found during a replace operation.
    public class MatchLogger : IReplacingCallback
    {
        public List<string> Matches { get; } = new List<string>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            Matches.Add(args.Match.Value);
            // Perform the default replacement.
            return ReplaceAction.Replace;
        }

        public string GetLog()
        {
            var sb = new StringBuilder();
            foreach (var match in Matches)
                sb.AppendLine($"Matched: \"{match}\"");
            return sb.ToString();
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths in the current directory.
            string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

            // -----------------------------------------------------------------
            // 1. Create a sample document.
            // -----------------------------------------------------------------
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Hello PLACEHOLDER.");
            builder.Writeln("This is a PLACEHOLDER test.");
            builder.Writeln("Another PLACEHOLDER appears here.");
            doc.Save(inputPath);

            // -----------------------------------------------------------------
            // 2. Load the document for processing.
            // -----------------------------------------------------------------
            var loadedDoc = new Document(inputPath);
            var reporter = new ReplacementProgressReporter();

            // -----------------------------------------------------------------
            // 3. First replacement: PLACEHOLDER -> First
            // -----------------------------------------------------------------
            var logger1 = new MatchLogger();
            var options1 = new FindReplaceOptions { ReplacingCallback = logger1 };
            int count1 = loadedDoc.Range.Replace("PLACEHOLDER", "First", options1);
            if (count1 == 0) throw new InvalidOperationException("Expected at least one replacement in step 1.");
            reporter.Report(count1);
            Console.WriteLine(logger1.GetLog());

            // -----------------------------------------------------------------
            // 4. Second replacement: First -> Second
            // -----------------------------------------------------------------
            var logger2 = new MatchLogger();
            var options2 = new FindReplaceOptions { ReplacingCallback = logger2 };
            int count2 = loadedDoc.Range.Replace("First", "Second", options2);
            if (count2 == 0) throw new InvalidOperationException("Expected at least one replacement in step 2.");
            reporter.Report(count2);
            Console.WriteLine(logger2.GetLog());

            // -----------------------------------------------------------------
            // 5. Third replacement: Second -> Final
            // -----------------------------------------------------------------
            var logger3 = new MatchLogger();
            var options3 = new FindReplaceOptions { ReplacingCallback = logger3 };
            int count3 = loadedDoc.Range.Replace("Second", "Final", options3);
            if (count3 == 0) throw new InvalidOperationException("Expected at least one replacement in step 3.");
            reporter.Report(count3);
            Console.WriteLine(logger3.GetLog());

            // -----------------------------------------------------------------
            // 6. Save the modified document.
            // -----------------------------------------------------------------
            loadedDoc.Save(outputPath);

            // Verify that the output file was created.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
