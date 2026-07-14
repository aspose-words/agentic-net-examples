using System;
using System.IO;
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
        builder.Writeln("This line contains foo but not at the start.");
        builder.Writeln("   foo with leading spaces should also be considered at start.");
        builder.Writeln("No occurrence here.");

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new StartOfParagraphReplacer()
        };

        // Replace the word "foo" with "bar" only when it appears at the start of a paragraph.
        int replacedCount = doc.Range.Replace("foo", "bar", options);

        // Validate that at least one replacement was performed.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }

    // Callback that allows replacement only when the match is at the start of its paragraph.
    private class StartOfParagraphReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match offset is the zero‑based position within the node that contains the match.
            // If the offset is zero, the match starts at the beginning of that node.
            // Additionally, ensure the node is the first text node of the paragraph.
            if (args.MatchOffset == 0)
            {
                // Verify that the node containing the match is the first child of its paragraph.
                var paragraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
                if (paragraph != null && paragraph.FirstChild == args.MatchNode)
                {
                    // Allow the replacement.
                    return ReplaceAction.Replace;
                }
            }

            // Skip replacement for all other occurrences.
            return ReplaceAction.Skip;
        }
    }
}
