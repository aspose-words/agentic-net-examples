using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with placeholders.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First line with {placeholder}.");
        builder.Writeln("Second line contains {placeholder} as well.");
        builder.Writeln("{placeholder} appears at the beginning of this line.");

        // Save the source document locally.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find-and-replace with a callback that inserts dynamic content after each replacement.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertAfterCallback()
        };

        // Perform the replacement.
        int replacedCount = loaded.Range.Replace("{placeholder}", "VALUE", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that inserts a line after the paragraph containing each match.
    private class InsertAfterCallback : IReplacingCallback
    {
        private int _matchIndex = 0;

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            _matchIndex++;

            // Keep the original replacement text.
            args.Replacement = "VALUE";

            // Locate the paragraph that contains the match.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Use a DocumentBuilder to insert a new line after the paragraph.
                // Cast the DocumentBase to Document to match the appropriate constructor.
                DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);
                cb.MoveTo(paragraph);
                cb.Writeln($"[Inserted after replacement #{_matchIndex}]");
            }

            return ReplaceAction.Replace;
        }
    }
}
