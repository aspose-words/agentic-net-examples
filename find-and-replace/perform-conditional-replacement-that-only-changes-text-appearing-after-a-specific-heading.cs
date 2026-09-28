using System;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Text before the heading (should NOT be replaced).
        builder.Writeln("placeholder before heading");

        // The specific heading after which replacements should occur.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");

        // Text after the heading (should be replaced).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("placeholder after heading");
        builder.Writeln("another placeholder after heading");

        // Save the input document (optional, just for demonstration).
        doc.Save("input.docx");

        // Load the document for processing.
        Document loaded = new Document("input.docx");

        // Set up find-and-replace with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ConditionalReplacer("Target Heading");

        // Perform the replacement: replace the word "placeholder" with "new value" only after the heading.
        int replacedCount = loaded.Range.Replace("placeholder", "new value", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the heading.");

        // Save the modified document.
        loaded.Save("output.docx");
    }
}

// Callback that replaces matches only if they appear after a specific heading.
public class ConditionalReplacer : IReplacingCallback
{
    private readonly string _headingText;

    public ConditionalReplacer(string headingText)
    {
        _headingText = headingText ?? throw new ArgumentNullException(nameof(headingText));
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Find the paragraph that contains the current match.
        Node matchNode = args.MatchNode;
        Paragraph currentParagraph = (Paragraph)matchNode.GetAncestor(NodeType.Paragraph);
        if (currentParagraph == null)
            return ReplaceAction.Skip;

        // Walk backwards through preceding sibling nodes to locate the heading.
        Node node = currentParagraph.PreviousSibling;
        while (node != null)
        {
            if (node.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)node;
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                    para.GetText().Trim() == _headingText)
                {
                    // Heading found before this match – allow replacement.
                    return ReplaceAction.Replace;
                }
            }
            node = node.PreviousSibling;
        }

        // No matching heading found before this match – skip replacement.
        return ReplaceAction.Skip;
    }
}
