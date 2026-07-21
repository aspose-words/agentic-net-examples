using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

public class ConditionalReplacementExample
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text before the heading – should NOT be replaced.
        builder.Writeln("Before heading placeholder.");
        builder.Writeln("PLACEHOLDER");

        // The specific heading after which replacements are allowed.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Start Here");
        builder.ParagraphFormat.ClearFormatting(); // Reset style for normal paragraphs.

        // Text after the heading – should be replaced.
        builder.Writeln("After heading placeholder.");
        builder.Writeln("PLACEHOLDER");

        // Locate the heading paragraph node.
        Paragraph headingParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
                                         .Cast<Paragraph>()
                                         .FirstOrDefault(p => p.GetText().Trim() == "Start Here");
        if (headingParagraph == null)
            throw new InvalidOperationException("Heading paragraph not found.");

        // Set up find/replace with a callback that checks the position of each match.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer(headingParagraph)
        };

        // Perform the replacement.
        int replacedCount = doc.Range.Replace("PLACEHOLDER", "REPLACED", options);

        // Verify that at least one replacement occurred after the heading.
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were performed.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Output result information.
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Modified document saved to: {Path.GetFullPath(outputPath)}");
    }

    // Callback that replaces only matches occurring after a specific heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly Paragraph _headingParagraph;
        private readonly int _headingIndex;

        public ConditionalReplacer(Paragraph headingParagraph)
        {
            _headingParagraph = headingParagraph ?? throw new ArgumentNullException(nameof(headingParagraph));

            // Determine the document order index of the heading paragraph.
            List<Node> allNodes = _headingParagraph.Document.GetChildNodes(NodeType.Any, true).Cast<Node>().ToList();
            _headingIndex = allNodes.IndexOf(_headingParagraph);
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the start of the match.
            Node current = args.MatchNode;
            while (current != null && current.NodeType != NodeType.Paragraph)
                current = current.ParentNode;

            Paragraph matchParagraph = current as Paragraph;
            if (matchParagraph == null)
                return ReplaceAction.Skip; // Safety fallback.

            // Determine the document order index of the match paragraph.
            List<Node> allNodes = matchParagraph.Document.GetChildNodes(NodeType.Any, true).Cast<Node>().ToList();
            int matchIndex = allNodes.IndexOf(matchParagraph);

            // Replace only if the match occurs after the heading.
            if (matchIndex > _headingIndex)
            {
                args.Replacement = "REPLACED";
                return ReplaceAction.Replace;
            }

            return ReplaceAction.Skip;
        }
    }
}
