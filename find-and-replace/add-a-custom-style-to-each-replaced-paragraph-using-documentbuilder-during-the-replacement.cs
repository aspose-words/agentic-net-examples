using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with paragraphs containing the target text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with placeholder.");
        builder.Writeln("Second paragraph with placeholder.");
        builder.Writeln("Third paragraph without the keyword.");

        // Save the initial document (optional, just to demonstrate file creation).
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        doc.Save(inputPath);

        // Define a custom paragraph style to apply after replacement.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "CustomReplacedStyle");
        customStyle.Font.Color = Color.Red; // Use System.Drawing.Color directly
        customStyle.Font.Bold = true;
        customStyle.Font.Size = 14;

        // Prepare the callback that will replace the text and apply the style.
        var callback = new ParagraphStyleReplacingCallback(customStyle);
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = callback
        };

        // Perform the find-and-replace operation.
        int replacedCount = doc.Range.Replace("placeholder", "UPDATED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Create a simple JSON report of the operation.
        var report = new
        {
            ReplacementsMade = replacedCount,
            ReplacedParagraphIndices = callback.ReplacedParagraphIndices
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "report.json");
        File.WriteAllText(reportPath, jsonReport);
    }

    // Callback that replaces matched text and applies a custom style to the containing paragraph.
    private class ParagraphStyleReplacingCallback : IReplacingCallback
    {
        private readonly Style _style;
        public List<int> ReplacedParagraphIndices { get; } = new List<int>();

        public ParagraphStyleReplacingCallback(Style style)
        {
            _style = style ?? throw new ArgumentNullException(nameof(style));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Replace the matched text with the desired new text.
            args.Replacement = "UPDATED";

            // Find the paragraph that contains the match.
            Node paragraphNode = args.MatchNode.GetAncestor(NodeType.Paragraph);
            if (paragraphNode is Paragraph paragraph)
            {
                // Apply the custom style.
                paragraph.ParagraphFormat.Style = _style;

                // Record the paragraph index for reporting.
                Body body = paragraph.GetAncestor(NodeType.Body) as Body;
                if (body != null)
                {
                    NodeCollection paragraphs = body.GetChildNodes(NodeType.Paragraph, true);
                    int index = paragraphs.IndexOf(paragraph);
                    if (index >= 0)
                        ReplacedParagraphIndices.Add(index);
                }
            }

            return ReplaceAction.Replace;
        }
    }
}
