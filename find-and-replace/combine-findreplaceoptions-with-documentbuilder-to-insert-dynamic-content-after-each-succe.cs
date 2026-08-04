using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First placeholder text.");
        builder.Writeln("Second placeholder appears here.");
        builder.Writeln("No match on this line.");

        // Save the sample document locally (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document that will be processed.
        Document loadedDoc = new Document(inputPath);

        // Configure find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertDynamicContentCallback()
        };

        // Replace the word "placeholder" with "value".
        int replacementCount = loadedDoc.Range.Replace("placeholder", "value", options);

        // Verify that at least one replacement was performed.
        if (replacementCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loadedDoc.Save(outputPath);
    }

    // Callback that inserts a new paragraph with dynamic content after each replacement.
    private class InsertDynamicContentCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The node that contains the start of the match.
            var matchNode = args.MatchNode;

            // Its parent paragraph.
            var paragraph = (Paragraph)matchNode.ParentNode;

            // Create a DocumentBuilder for the same document.
            // The Document property returns DocumentBase, so cast to Document.
            var builder = new DocumentBuilder((Document)matchNode.Document);

            // Move the builder to the paragraph that contains the match.
            builder.MoveTo(paragraph);

            // Insert a new paragraph after the current one.
            builder.InsertParagraph();

            // Write dynamic content into the newly inserted paragraph.
            builder.Writeln($"[Inserted after replacement at {DateTime.Now:HH:mm:ss}]");

            // Continue with the normal replacement.
            return ReplaceAction.Replace;
        }
    }
}
