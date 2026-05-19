using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two headings and several occurrences of the word "target".
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Heading 1");
        builder.Writeln("This is a target before the second heading.");
        builder.Writeln("Heading 2");
        builder.Writeln("First target after heading 2.");
        builder.Writeln("Another target after heading 2.");

        // Set up find‑replace options with a custom callback that only replaces matches
        // that appear after the specified heading.
        var options = new FindReplaceOptions
        {
            ReplacingCallback = new ConditionalReplacer("Heading 2", "REPLACED")
        };

        // Perform the replacement. The callback decides whether each match should be replaced.
        int replacedCount = doc.Range.Replace("target", "REPLACED", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No replacements were made; the conditional logic may be incorrect.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Optional: indicate success (no interactive prompts required).
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }

    // Callback that replaces a match only if it is located after a paragraph containing a specific heading.
    private class ConditionalReplacer : IReplacingCallback
    {
        private readonly string _headingText;
        private readonly string _replacementText;

        public ConditionalReplacer(string headingText, string replacementText)
        {
            _headingText = headingText ?? throw new ArgumentNullException(nameof(headingText));
            _replacementText = replacementText ?? throw new ArgumentNullException(nameof(replacementText));
        }

        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The node that contains the beginning of the match (usually a Run).
            Node matchNode = args.MatchNode;
            // Ascend to the containing paragraph.
            Paragraph paragraph = matchNode?.ParentNode as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Walk backwards through preceding sibling nodes to see if the heading appears.
            Node current = paragraph.PreviousSibling;
            while (current != null)
            {
                if (current.NodeType == NodeType.Paragraph)
                {
                    string text = current.GetText().Trim();
                    if (text.Equals(_headingText, StringComparison.OrdinalIgnoreCase))
                    {
                        // Heading found before this paragraph – perform replacement.
                        args.Replacement = _replacementText;
                        return ReplaceAction.Replace;
                    }

                    // If another heading is encountered before the target heading, stop searching.
                    // Assuming headings are distinct lines; adjust as needed.
                    if (text.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                        break;
                }
                current = current.PreviousSibling;
            }

            // Heading not found in preceding paragraphs – skip this match.
            return ReplaceAction.Skip;
        }
    }
}
