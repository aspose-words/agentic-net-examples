using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with headings and body text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading containing the keyword.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Special Report Overview");

        // Heading without the keyword.
        builder.Writeln("General Summary");

        // Normal paragraph containing the keyword (should not be replaced).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This paragraph mentions Special but is not a heading.");

        // Another heading containing the keyword.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Special Findings");

        // Save the input document.
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up find-and-replace options with a callback that limits replacements to headings.
        FindReplaceOptions options = new FindReplaceOptions();
        options.ReplacingCallback = new HeadingReplaceCallback();

        // Perform the replacement: replace the word "Special" with "Replaced" only in headings.
        int replacedCount = loaded.Range.Replace("Special", "Replaced", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement in headings, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that allows replacement only when the match is inside a heading paragraph.
    private class HeadingReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match node is typically a Run.
            if (args.MatchNode is Run run)
            {
                Paragraph paragraph = run.ParentParagraph;
                if (paragraph != null)
                {
                    // Check if the paragraph style is a heading style.
                    StyleIdentifier styleId = paragraph.ParagraphFormat.StyleIdentifier;
                    if (styleId == StyleIdentifier.Heading1 ||
                        styleId == StyleIdentifier.Heading2 ||
                        styleId == StyleIdentifier.Heading3 ||
                        styleId == StyleIdentifier.Heading4 ||
                        styleId == StyleIdentifier.Heading5 ||
                        styleId == StyleIdentifier.Heading6 ||
                        styleId == StyleIdentifier.Heading7 ||
                        styleId == StyleIdentifier.Heading8 ||
                        styleId == StyleIdentifier.Heading9)
                    {
                        // Allow the replacement.
                        return ReplaceAction.Replace;
                    }
                }
            }

            // Skip replacement for all other cases.
            return ReplaceAction.Skip;
        }
    }
}
