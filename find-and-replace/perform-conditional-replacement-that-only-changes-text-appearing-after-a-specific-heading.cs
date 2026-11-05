using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text before the heading – should NOT be replaced.
        builder.Writeln("Intro paragraph with placeholder.");

        // Insert the target heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");

        // Text after the heading – should be replaced.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph contains a placeholder that will be changed.");

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Locate the heading paragraph.
        Paragraph headingParagraph = null;
        foreach (Paragraph para in loaded.GetChildNodes(NodeType.Paragraph, true))
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                para.GetText().Trim() == "Target Heading")
            {
                headingParagraph = para;
                break;
            }
        }

        if (headingParagraph == null)
            throw new InvalidOperationException("Heading not found in the document.");

        // Set up the find‑replace with a callback that only replaces after the heading.
        var options = new FindReplaceOptions();
        options.ReplacingCallback = new ConditionalReplacer(loaded, headingParagraph);

        int replacedCount = loaded.Range.Replace("placeholder", "replaced", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the heading.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that skips replacements occurring before the specified heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly int _headingParagraphIndex;
        private readonly List<Paragraph> _allParagraphs;

        public ConditionalReplacer(Document document, Paragraph headingParagraph)
        {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (headingParagraph == null) throw new ArgumentNullException(nameof(headingParagraph));

            // Cache all paragraphs for index lookup.
            _allParagraphs = new List<Paragraph>();
            foreach (Paragraph para in document.GetChildNodes(NodeType.Paragraph, true))
                _allParagraphs.Add(para);

            _headingParagraphIndex = _allParagraphs.IndexOf(headingParagraph);
            if (_headingParagraphIndex < 0)
                throw new InvalidOperationException("Heading paragraph not found in cached list.");
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the current match.
            Paragraph matchParagraph = args.MatchNode?.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (matchParagraph == null)
                return ReplaceAction.Skip;

            int matchIndex = _allParagraphs.IndexOf(matchParagraph);
            // If the match paragraph appears after the heading, replace; otherwise skip.
            return matchIndex > _headingParagraphIndex ? ReplaceAction.Replace : ReplaceAction.Skip;
        }
    }
}
