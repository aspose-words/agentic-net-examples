using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a new document and populate it with headings and sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First heading and some text before the target heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a target before heading.");
        builder.Writeln("target");

        // Second heading – replacements should start after this heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Section 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("target should be replaced");
        builder.Writeln("Another target.");

        // Locate the heading paragraph after which replacements are allowed.
        Paragraph headingParagraph = doc.GetChildNodes(NodeType.Paragraph, true)
            .Cast<Paragraph>()
            .FirstOrDefault(p =>
                p.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                p.GetText().Trim().Equals("Section 2", StringComparison.Ordinal));

        // Set up the find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer(headingParagraph)
        };

        // Replace the word "target" only after the specified heading.
        int replacedCount = doc.Range.Replace("target", "replaced", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the heading.");

        // Save the modified document.
        doc.Save("output.docx");

        // Create a simple JSON report of the operation.
        var report = new { Replacements = replacedCount };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("report.json", json);
    }

    // Callback that performs replacement only after a given heading paragraph.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly Body _body;
        private readonly int _headingParagraphIndex;

        public ConditionalReplacer(Paragraph headingParagraph)
        {
            if (headingParagraph?.ParentNode is Body body)
            {
                _body = body;
                _headingParagraphIndex = body.Paragraphs.IndexOf(headingParagraph);
            }
            else
            {
                _body = null;
                _headingParagraphIndex = -1;
            }
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Identify the paragraph that contains the current match.
            Paragraph matchParagraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (matchParagraph == null)
                return ReplaceAction.Skip;

            // Ensure we are working within the same body as the heading.
            if (matchParagraph.ParentNode is Body body && body == _body)
            {
                int matchIndex = body.Paragraphs.IndexOf(matchParagraph);
                if (matchIndex > _headingParagraphIndex)
                {
                    args.Replacement = "replaced";
                    return ReplaceAction.Replace;
                }
            }

            // Skip matches that appear before the heading.
            return ReplaceAction.Skip;
        }
    }
}
