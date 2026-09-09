using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Product: {Item}");
        builder.Writeln("Price: {Item}");
        builder.Writeln("Description: {Item}");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertAfterReplacementCallback()
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("{Item}", "Widget", options);

        // Ensure that at least one replacement was made.
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that inserts a new paragraph after each replacement.
    private class InsertAfterReplacementCallback : IReplacingCallback
    {
        private int _matchIndex = 0;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            _matchIndex++;

            // Set the replacement text.
            args.Replacement = "Widget";

            // Find the paragraph that contains the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Use a DocumentBuilder positioned at the found paragraph.
                DocumentBuilder cb = new DocumentBuilder((Document)paragraph.Document);
                cb.MoveTo(paragraph);
                // Insert a new paragraph after the current one.
                cb.InsertParagraph();
                cb.Writeln($"[Inserted after replacement #{_matchIndex}]");
            }

            // Continue with the replacement.
            return ReplaceAction.Replace;
        }
    }
}
