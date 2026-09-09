using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with various occurrences of the word "foo".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("foo appears at the start of this paragraph.");
        builder.Writeln("This paragraph contains foo in the middle.");
        builder.Writeln("Another line with foo at the start.");
        builder.Writeln("No match here.");
        builder.Writeln("foo"); // paragraph that only contains the word

        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new StartOfParagraphReplacer();

        // Replace the word "foo" with "bar" only when it is at the start of a paragraph.
        int replacedCount = loaded.Range.Replace("foo", "bar", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        loaded.Save(outputPath);
    }

    // Callback that replaces only matches that start a paragraph.
    private class StartOfParagraphReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the paragraph that contains the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Find the first Run node in the paragraph.
            Run firstRun = paragraph.GetChildNodes(NodeType.Run, true)[0] as Run;
            if (firstRun == null)
                return ReplaceAction.Skip;

            // The match must start at the very beginning of the first Run.
            bool isAtParagraphStart = args.MatchNode == firstRun && args.MatchOffset == 0;

            return isAtParagraphStart ? ReplaceAction.Replace : ReplaceAction.Skip;
        }
    }
}
