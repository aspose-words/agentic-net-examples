using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;          // Required package reference
using Newtonsoft.Json;        // Required package reference

public class Program
{
    public static void Main()
    {
        // Prepare a folder for all temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document with text that will be replaced.
        // -----------------------------------------------------------------
        string inputPath = Path.Combine(workDir, "input.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("alpha beta alpha gamma alpha");
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and configure the replacement callback.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);
        var logger = new ReplaceLogger();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = logger
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("alpha", "omega", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // -----------------------------------------------------------------
        // 3. Save the modified document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(workDir, "output.docx");
        loaded.Save(outputPath);

        // -----------------------------------------------------------------
        // 4. Write the log of replacements to a text file.
        // -----------------------------------------------------------------
        string logPath = Path.Combine(workDir, "replace_log.txt");
        File.WriteAllLines(logPath, logger.Matches);

        // Validate that the log file was created.
        if (!File.Exists(logPath))
            throw new FileNotFoundException("Log file was not created.", logPath);
    }

    // -----------------------------------------------------------------
    // Custom logger that records each match found during replacement.
    // -----------------------------------------------------------------
    private class ReplaceLogger : IReplacingCallback
    {
        public List<string> Matches { get; } = new List<string>();

        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
        {
            // Record the original matched text.
            Matches.Add(args.Match.Value);
            // Proceed with the replacement.
            return ReplaceAction.Replace;
        }
    }
}
