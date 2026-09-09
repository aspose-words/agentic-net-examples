using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a blank document and add some markdown‑style headings.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("# Document Title");
        builder.Writeln("Some introductory text.");
        builder.Writeln("## Chapter 1");
        builder.Writeln("Content of chapter 1.");
        builder.Writeln("### Section 1.1");
        builder.Writeln("Details of section 1.1.");
        builder.Writeln("## Chapter 2");
        builder.Writeln("Content of chapter 2.");
        builder.Writeln("Normal paragraph without heading.");

        // Regular expression that matches markdown headings (levels 1‑6).
        Regex headingRegex = new Regex(@"^(#{1,6})\s+(.*)$", RegexOptions.Multiline);

        // Set up replace options with a callback that removes the markdown symbols
        // and applies the corresponding Word heading style.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new MarkdownHeadingReplacer()
        };

        int replacedCount = doc.Range.Replace(headingRegex, "$2", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("No markdown headings were found for replacement.");

        // Save the modified document.
        doc.Save("output.docx");
    }

    // Callback that formats each matched heading.
    private class MarkdownHeadingReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Determine heading level from the number of leading '#'.
            int level = args.Match.Groups[1].Value.Length; // 1‑6

            // Extract the plain heading text (without markdown symbols).
            string headingText = args.Match.Groups[2].Value;
            args.Replacement = headingText; // Replace the whole match with plain text.

            // The match node is usually a Run; its parent is the Paragraph that holds the heading.
            if (args.MatchNode?.ParentNode is Paragraph paragraph)
            {
                // Apply the appropriate built‑in heading style.
                // StyleIdentifier.Heading1, Heading2, … are sequential enum values.
                paragraph.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1 + (level - 1);
            }

            return ReplaceAction.Replace;
        }
    }
}
