using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a heading and placeholders before and after it.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text before the heading (should NOT be replaced).
        builder.Writeln("Intro paragraph with a placeholder:");
        builder.Writeln("PLACEHOLDER");

        // The target heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Text after the heading (should be replaced).
        builder.Writeln("Paragraph after heading with a placeholder:");
        builder.Writeln("PLACEHOLDER");

        // Save the source document.
        doc.Save(inputPath);

        // Reload the document to simulate a real‑world scenario.
        Document loaded = new Document(inputPath);

        // Locate the heading paragraph that marks the start of the replacement region.
        Paragraph headingParagraph = loaded.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .FirstOrDefault(p => p.GetText().Trim() == "Target Heading");

        if (headingParagraph == null)
            throw new InvalidOperationException("Heading not found in the document.");

        // Set up the find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer(headingParagraph)
        };

        // Perform the replacement; only matches after the heading will be changed.
        int replacedCount = loaded.Range.Replace("PLACEHOLDER", "REPLACED", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the heading.");

        // Save the modified document.
        loaded.Save(outputPath);

        // Output the result count (optional, just to demonstrate execution).
        Console.WriteLine($"Replacements performed: {replacedCount}");
    }

    // Callback that replaces matches only if they appear after a specific heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly Paragraph _headingParagraph;
        private readonly Paragraph[] _allParagraphs;

        public ConditionalReplacer(Paragraph headingParagraph)
        {
            _headingParagraph = headingParagraph ?? throw new ArgumentNullException(nameof(headingParagraph));
            // Cache the ordered list of all paragraphs for index comparison.
            _allParagraphs = headingParagraph.Document
                .GetChildNodes(NodeType.Paragraph, true)
                .Cast<Paragraph>()
                .ToArray();
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine the paragraph that contains the current match.
            Paragraph matchParagraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (matchParagraph == null)
                return ReplaceAction.Skip;

            // Compare the positions of the match paragraph and the heading paragraph.
            int headingIndex = Array.IndexOf(_allParagraphs, _headingParagraph);
            int matchIndex = Array.IndexOf(_allParagraphs, matchParagraph);

            // Replace only if the match occurs after the heading.
            if (matchIndex > headingIndex)
            {
                args.Replacement = "REPLACED";
                return ReplaceAction.Replace;
            }

            return ReplaceAction.Skip;
        }
    }
}
