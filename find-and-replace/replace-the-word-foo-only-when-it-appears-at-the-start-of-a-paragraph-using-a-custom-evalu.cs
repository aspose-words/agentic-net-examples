using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("foo is at the start of this paragraph.");
        builder.Writeln("This foo is not at the start.");
        builder.Writeln("foo");
        builder.Writeln("Another line with foo at the start.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new StartOfParagraphCallback()
        };

        // Perform the replace operation.
        int replacedCount = doc.Range.Replace("foo", "bar", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that replaces only matches that appear at the start of a paragraph.
    private class StartOfParagraphCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The node that contains the beginning of the match.
            Node matchNode = args.MatchNode;

            // Its parent should be a paragraph.
            if (matchNode?.ParentNode is not Paragraph paragraph)
                return ReplaceAction.Skip;

            // Determine if the match is at the very beginning of the paragraph.
            bool isAtParagraphStart = args.MatchOffset == 0 && matchNode == paragraph.FirstChild;

            if (isAtParagraphStart)
            {
                // Replace the word.
                args.Replacement = "bar";
                return ReplaceAction.Replace;
            }

            // Otherwise, skip this match.
            return ReplaceAction.Skip;
        }
    }
}
