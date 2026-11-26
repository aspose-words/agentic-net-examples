using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("foo is at the start of this paragraph.");
        builder.Writeln("This paragraph contains foo but not at the start.");
        builder.Writeln("foo");
        builder.Writeln("Another line without the keyword.");

        // Save the input document (optional, just for demonstration).
        const string inputPath = "input.docx";
        doc.Save(inputPath);

        // Prepare the find‑replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new StartParagraphCallback()
        };

        // Use a regex to find the word "foo" as a whole word.
        Regex regex = new Regex(@"\bfoo\b", RegexOptions.None);

        // Perform the replacement, changing "foo" to "bar" only at paragraph starts.
        int replacedCount = doc.Range.Replace(regex, "bar", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }

    // Custom callback that replaces only when the match is at the start of a paragraph.
    private class StartParagraphCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Ensure the match node is a Run.
            if (args.MatchNode is Run runNode)
            {
                // Get the paragraph that contains this run.
                Paragraph paragraph = runNode.ParentParagraph;

                // Guard against nulls (should not happen, but defensive programming).
                if (paragraph == null || args.Match == null)
                    return ReplaceAction.Skip;

                // Determine if the paragraph text starts with the matched word.
                // Paragraph.GetText() includes the paragraph mark at the end, so we trim it.
                string paragraphText = paragraph.GetText().TrimEnd('\r', '\a');

                // If the paragraph starts exactly with the matched value, allow replacement.
                if (paragraphText.StartsWith(args.Match.Value, StringComparison.Ordinal))
                    return ReplaceAction.Replace;
            }

            // Skip replacement for all other cases.
            return ReplaceAction.Skip;
        }
    }
}
