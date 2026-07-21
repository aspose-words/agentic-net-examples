using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "report.json");

        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line with PLACEHOLDER.");
        builder.Writeln("Second line without.");
        builder.Writeln("Third line with another PLACEHOLDER.");
        doc.Save(inputPath);

        // Load the document (could also continue using the same instance).
        Document loadedDoc = new Document(inputPath);

        // Set up the replace callback.
        var callback = new InsertAfterReplaceCallback();

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the replacement.
        int replacedCount = loadedDoc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loadedDoc.Save(outputPath);

        // Write a JSON report of all matches.
        string json = JsonConvert.SerializeObject(callback.MatchedValues, Formatting.Indented);
        File.WriteAllText(reportPath, json);

        // Validate that the report file was created.
        if (!File.Exists(reportPath))
            throw new InvalidOperationException("Report file was not created.");

        // Example completed successfully (no interactive output).
    }

    // Callback that records matches and inserts dynamic content after each replacement.
    private class InsertAfterReplaceCallback : IReplacingCallback
    {
        public List<string> MatchedValues { get; } = new List<string>();

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Record the matched text.
            if (args.Match?.Value != null)
                MatchedValues.Add(args.Match.Value);

            // Insert dynamic content after the matched node.
            if (args.MatchNode?.Document is Document doc && args.MatchNode != null)
            {
                // Find the paragraph that contains the match node.
                Paragraph? paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
                if (paragraph != null && paragraph.ParentNode != null)
                {
                    // Create a new paragraph with the dynamic content.
                    Paragraph newParagraph = new Paragraph(doc);
                    paragraph.ParentNode.InsertAfter(newParagraph, paragraph);

                    // Write the dynamic content into the new paragraph.
                    DocumentBuilder cb = new DocumentBuilder(doc);
                    cb.MoveTo(newParagraph);
                    cb.Writeln($"[Replaced at {DateTime.Now:O}]");
                }
            }

            // Proceed with the standard replacement.
            return ReplaceAction.Replace;
        }
    }
}
