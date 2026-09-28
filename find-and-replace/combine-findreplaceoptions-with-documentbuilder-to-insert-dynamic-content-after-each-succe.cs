using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class InsertAfterCallback : IReplacingCallback
{
    private readonly string _dynamicPrefix;
    private int _counter = 0;
    public List<string> MatchedValues { get; } = new List<string>();

    public InsertAfterCallback(string dynamicPrefix)
    {
        _dynamicPrefix = dynamicPrefix;
    }

    ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
    {
        // Record the matched text.
        MatchedValues.Add(args.Match.Value);

        // Increment counter to generate unique dynamic content.
        _counter++;

        // Use a DocumentBuilder tied to the document that contains the match.
        DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);

        // Move the builder to the node that contains the match.
        builder.MoveTo(args.MatchNode);

        // Insert a new paragraph after the current node.
        Paragraph insertedParagraph = builder.InsertParagraph();

        // Write the dynamic content into the newly inserted paragraph.
        builder.Writeln($"{_dynamicPrefix} {_counter}");

        // Proceed with the normal replacement.
        return ReplaceAction.Replace;
    }
}

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a PLACEHOLDER that will be replaced.");
        builder.Writeln("Another PLACEHOLDER appears here.");
        builder.Writeln("No placeholder on this line.");

        // Save the input document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loadedDoc = new Document(inputPath);

        // Set up the callback to insert dynamic content after each replacement.
        InsertAfterCallback callback = new InsertAfterCallback("Inserted after replacement");
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the find-and-replace operation.
        int replacedCount = loadedDoc.Range.Replace("PLACEHOLDER", "REPLACED", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none occurred.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Serialize the list of matched values to a JSON report.
        string jsonReport = JsonConvert.SerializeObject(callback.MatchedValues, Formatting.Indented);
        const string reportPath = "matches.json";
        File.WriteAllText(reportPath, jsonReport);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Failed to create the matches report file.");
    }
}
