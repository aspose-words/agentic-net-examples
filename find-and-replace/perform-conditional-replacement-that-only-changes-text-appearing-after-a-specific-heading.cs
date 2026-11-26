using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Intro paragraph before heading.");

        // First heading – the target heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Target Heading");

        // Content after the target heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This text should be replaced: placeholder.");
        builder.Writeln("Another placeholder line.");

        // Another heading – not the target.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Another Heading");
        builder.Writeln("placeholder after second heading.");

        // Set up find‑replace with a callback that only replaces after the target heading.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new ConditionalReplacer("Target Heading");

        int replacedCount = doc.Range.Replace("placeholder", "REPLACED", options);

        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement after the target heading.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Replacements performed: {replacedCount}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Callback that allows replacement only if the match appears after a specific heading.
public class ConditionalReplacer : IReplacingCallback
{
    private readonly string _targetHeadingText;

    public ConditionalReplacer(string targetHeadingText)
    {
        _targetHeadingText = targetHeadingText;
    }

    public ReplaceAction Replacing(ReplacingArgs args)
    {
        // Find the paragraph that contains the current match.
        Paragraph matchParagraph = (Paragraph)args.MatchNode.GetAncestor(NodeType.Paragraph);
        if (matchParagraph == null)
            return ReplaceAction.Skip;

        // Walk backwards through previous siblings to locate the target heading.
        Node current = matchParagraph.PreviousSibling;
        while (current != null)
        {
            if (current.NodeType == NodeType.Paragraph)
            {
                Paragraph para = (Paragraph)current;
                if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1 &&
                    para.GetText().Trim() == _targetHeadingText)
                {
                    // The target heading is found before this match – allow replacement.
                    return ReplaceAction.Replace;
                }
            }
            current = current.PreviousSibling;
        }

        // No target heading before this match – skip replacement.
        return ReplaceAction.Skip;
    }
}
