using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with multiple placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a PLACEHOLDER text.");
        builder.Writeln("Another PLACEHOLDER appears here.");
        builder.Writeln("PLACEHOLDER at the start of a line.");
        builder.Writeln("No placeholder on this line.");

        // Save the input document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Callback to count replacements.
        var replaceCounter = new ReplaceCounter();

        // Set up FindReplaceOptions with the callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = replaceCounter
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Insert dynamic content after each successful replacement using DocumentBuilder logic.
        int insertionIndex = 1;
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "REPLACED")
            {
                // Create a new run with dynamic content.
                Run insertedRun = new Run(doc, $"[Inserted {insertionIndex}]");
                // Insert the new run after the replaced run.
                run.ParentNode.InsertAfter(insertedRun, run);
                insertionIndex++;
            }
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Create a simple JSON report of the operation.
        var report = new
        {
            OriginalPlaceholdersFound = replaceCounter.MatchCount,
            PlaceholdersReplaced = replacedCount,
            DynamicInsertionsMade = insertionIndex - 1,
            InputDocument = Path.GetFullPath(inputPath),
            OutputDocument = Path.GetFullPath(outputPath)
        };

        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        const string reportPath = "report.json";
        File.WriteAllText(reportPath, jsonReport);
    }

    // Callback implementation that counts matches.
    private class ReplaceCounter : IReplacingCallback
    {
        public int MatchCount { get; private set; }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            MatchCount++;
            // Use the default replacement behavior.
            return ReplaceAction.Replace;
        }
    }
}
