using System;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample paragraphs.
        builder.Writeln("foo is at the start of this paragraph.");
        builder.Writeln("This paragraph contains foo but not at the start.");
        builder.Writeln("foo appears again at the beginning.");
        builder.Writeln("No matching word here.");

        // Set up find/replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions(new FooStartCallback());

        // Replace the word "foo" with "bar" only when it is at the start of a paragraph.
        int replacementCount = doc.Range.Replace("foo", "bar", options);

        // Validate that at least one replacement was made.
        if (replacementCount == 0)
            throw new InvalidOperationException("No occurrences of 'foo' at the start of a paragraph were found.");

        // Save the modified document.
        const string outputPath = "Modified.docx";
        doc.Save(outputPath);

        // Optional: output the number of replacements performed.
        Console.WriteLine($"Replacements made: {replacementCount}");
    }

    // Custom callback that replaces only when the match is at the start of a paragraph.
    private class FooStartCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the beginning of the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Check if the paragraph text starts with the matched word.
            string paragraphText = paragraph.GetText(); // Includes paragraph break at the end.
            if (paragraphText.StartsWith(args.Match.Value, StringComparison.Ordinal))
            {
                // Perform the replacement.
                args.Replacement = "bar";
                return ReplaceAction.Replace;
            }

            // Skip this match because it is not at the start of the paragraph.
            return ReplaceAction.Skip;
        }
    }
}
