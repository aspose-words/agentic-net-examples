using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

namespace FindAndReplaceProgressDemo
{
    // Simple progress reporter that receives the replacement count after each Replace call.
    public class ProgressReporter
    {
        public void Report(int replacementCount, string operationDescription)
        {
            Console.WriteLine($"{operationDescription}: {replacementCount} replacement(s) performed.");
        }
    }

    // Optional logger that records each match found during a replace operation.
    public class ReplaceLogger : IReplacingCallback
    {
        public List<string> Matches { get; } = new List<string>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            Matches.Add(args.Match.Value);
            // Perform the default replacement.
            return ReplaceAction.Replace;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Alpha beta gamma. Alpha beta gamma. Alpha beta gamma.");
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the document from the file system.
            Document loadedDoc = new Document(inputPath);

            // Initialize progress reporter.
            ProgressReporter reporter = new ProgressReporter();

            // Optional logger to demonstrate callback usage.
            ReplaceLogger logger = new ReplaceLogger();
            FindReplaceOptions options = new FindReplaceOptions
            {
                ReplacingCallback = logger
            };

            // First replacement: plain text.
            int countAlpha = loadedDoc.Range.Replace("Alpha", "Delta", options);
            reporter.Report(countAlpha, "Replace \"Alpha\" with \"Delta\"");

            // Second replacement: plain text.
            int countBeta = loadedDoc.Range.Replace("beta", "Epsilon", options);
            reporter.Report(countBeta, "Replace \"beta\" with \"Epsilon\"");

            // Third replacement: regex pattern.
            int countGamma = loadedDoc.Range.Replace(new Regex("gamma"), "Zeta", options);
            reporter.Report(countGamma, "Replace \"gamma\" with \"Zeta\"");

            // Validate that at least one replacement occurred.
            if (countAlpha + countBeta + countGamma == 0)
                throw new InvalidOperationException("No replacements were performed.");

            // Save the modified document.
            const string outputPath = "output.docx";
            loadedDoc.Save(outputPath);

            // Optional: write a simple JSON report of matches (demonstrates Newtonsoft.Json usage).
            var report = new
            {
                Replacements = new[]
                {
                    new { From = "Alpha", To = "Delta", Count = countAlpha },
                    new { From = "beta", To = "Epsilon", Count = countBeta },
                    new { From = "gamma", To = "Zeta", Count = countGamma }
                },
                MatchesFound = logger.Matches
            };
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(report, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText("replace_report.json", json);
        }
    }
}
