using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("foo starts this paragraph.");
        builder.Writeln("This paragraph contains foo but not at the start.");
        builder.Writeln("foo");
        builder.Writeln("Another line without the keyword.");
        builder.Writeln("foobar should not be replaced because it is not a whole word at start.");
        builder.Writeln("   foo with leading spaces should not be replaced.");

        // Save the source document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new StartOfParagraphCallback());

        // Replace the word "foo" only when it appears at the start of a paragraph.
        int replacedCount = loaded.Range.Replace("foo", "bar", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that replaces only matches that start at the beginning of a paragraph.
    private class StartOfParagraphCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // args.MatchOffset is the zero‑based position of the match within the node that contains the start.
            // If the offset is zero, the match begins at the start of that node.
            // Additionally, ensure the match is at the very start of the paragraph.
            if (args.MatchOffset == 0 && IsFirstNodeInParagraph(args.MatchNode))
            {
                // Perform the replacement.
                args.Replacement = "bar";
                return ReplaceAction.Replace;
            }

            // Skip replacement for all other occurrences.
            return ReplaceAction.Skip;
        }

        // Determines whether the node containing the match is the first visible node of its paragraph.
        private static bool IsFirstNodeInParagraph(Node matchNode)
        {
            // Ascend to the containing paragraph.
            Paragraph paragraph = matchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph == null)
                return false;

            // Find the first child node of the paragraph that contains text.
            foreach (Node child in paragraph.GetChildNodes(NodeType.Any, true))
            {
                // Skip empty runs or nodes without text.
                if (child is Run run && string.IsNullOrEmpty(run.Text))
                    continue;

                // The first non‑empty node is the start of the paragraph.
                return child == matchNode;
            }

            return false;
        }
    }
}
