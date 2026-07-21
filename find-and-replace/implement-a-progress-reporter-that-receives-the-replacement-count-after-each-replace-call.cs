using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Prepare temporary file paths.
        string inputPath = Path.Combine(Path.GetTempPath(), "find_replace_input.docx");
        string outputPath = Path.Combine(Path.GetTempPath(), "find_replace_output.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma beta delta alpha.");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document for processing.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        // -----------------------------------------------------------------
        // 3. Set up a progress reporter.
        // -----------------------------------------------------------------
        var reporter = new ProgressReporter();

        // -----------------------------------------------------------------
        // 4. First replacement: "alpha" -> "ALPHA".
        // -----------------------------------------------------------------
        var logger1 = new ReplacementLogger();
        var options1 = new FindReplaceOptions { ReplacingCallback = logger1 };
        int replaced1 = loaded.Range.Replace("alpha", "ALPHA", options1);
        reporter.Report(replaced1); // Report after first replace.

        // -----------------------------------------------------------------
        // 5. Second replacement: "beta" -> "BETA".
        // -----------------------------------------------------------------
        var logger2 = new ReplacementLogger();
        var options2 = new FindReplaceOptions { ReplacingCallback = logger2 };
        int replaced2 = loaded.Range.Replace("beta", "BETA", options2);
        reporter.Report(replaced2); // Report after second replace.

        // -----------------------------------------------------------------
        // 6. Third replacement using a regex: "gamma|delta" -> "GAMMA_DELTA".
        // -----------------------------------------------------------------
        var logger3 = new ReplacementLogger();
        var options3 = new FindReplaceOptions { ReplacingCallback = logger3 };
        int replaced3 = loaded.Range.Replace(new System.Text.RegularExpressions.Regex(@"gamma|delta"), "GAMMA_DELTA", options3);
        reporter.Report(replaced3); // Report after third replace.

        // Validate that at least one replacement occurred overall.
        if (replaced1 + replaced2 + replaced3 == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // -----------------------------------------------------------------
        // 7. Save the modified document.
        // -----------------------------------------------------------------
        loaded.Save(outputPath);

        // -----------------------------------------------------------------
        // 8. Output final verification (optional, not interactive).
        // -----------------------------------------------------------------
        Console.WriteLine($"Processing complete. Output saved to: {outputPath}");
    }
}

// ---------------------------------------------------------------------
// Helper class that logs each match found during a replace operation.
// ---------------------------------------------------------------------
public class ReplacementLogger : IReplacingCallback
{
    public List<string> Matches { get; } = new List<string>();

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        Matches.Add(args.Match.Value);
        // Perform the default replacement.
        return ReplaceAction.Replace;
    }
}

// ---------------------------------------------------------------------
// Simple progress reporter that receives the replacement count.
// ---------------------------------------------------------------------
public class ProgressReporter
{
    public void Report(int replacementCount)
    {
        Console.WriteLine($"Replacements made in this step: {replacementCount}");
    }
}
