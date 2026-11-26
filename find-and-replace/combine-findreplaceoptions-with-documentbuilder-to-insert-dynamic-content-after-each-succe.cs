using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class InsertAfterCallback : IReplacingCallback
{
    public List<ReplacementRecord> Records { get; } = new List<ReplacementRecord>();

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Store original matched text.
        string originalText = args.Match.Value;

        // Define the replacement text.
        string replacementText = "REPLACED";

        // Set the replacement.
        args.Replacement = replacementText;

        // Insert dynamic content after the replaced node.
        // Cast DocumentBase to Document for the builder constructor.
        DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
        // Move the builder to the node that was replaced.
        builder.MoveTo(args.MatchNode);
        // Insert a new paragraph after the current node.
        builder.InsertParagraph();
        // Write dynamic content (e.g., timestamp) into the newly inserted paragraph.
        string dynamicContent = $"[Inserted at {DateTime.Now:O}]";
        builder.Writeln(dynamicContent);

        // Record the operation for reporting.
        Records.Add(new ReplacementRecord
        {
            OriginalText = originalText,
            ReplacementText = replacementText,
            InsertedContent = dynamicContent
        });

        // Continue with the replacement.
        return ReplaceAction.Replace;
    }
}

public class ReplacementRecord
{
    public string OriginalText { get; set; } = string.Empty;
    public string ReplacementText { get; set; } = string.Empty;
    public string InsertedContent { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("Find and replace the word PLACEHOLDER multiple times.");
        builder.Writeln("PLACEHOLDER appears here.");
        builder.Writeln("Another PLACEHOLDER is here.");

        // Save the initial document (optional, for inspection).
        doc.Save("input.docx");

        // Load the document for processing.
        Document loadedDoc = new Document("input.docx");

        // Set up the find-replace options with a callback.
        InsertAfterCallback callback = new InsertAfterCallback();
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback,
            MatchCase = false,
            FindWholeWordsOnly = true
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loadedDoc.Save("output.docx");

        // Serialize replacement records to JSON.
        string jsonReport = JsonConvert.SerializeObject(callback.Records, Formatting.Indented);
        File.WriteAllText("report.json", jsonReport);

        // Validate that the report file was created.
        if (!File.Exists("report.json"))
            throw new InvalidOperationException("Failed to create the JSON report file.");
    }
}
