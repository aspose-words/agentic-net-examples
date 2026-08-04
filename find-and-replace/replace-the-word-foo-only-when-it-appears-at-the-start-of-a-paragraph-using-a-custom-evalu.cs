using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Replacing;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Create a sample document with several paragraphs.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("foo is at the start of this paragraph.");          // Should be replaced.
        builder.Writeln("This line contains foo but not at the start.");   // Should stay unchanged.
        builder.Writeln("foo appears again at the beginning.");            // Should be replaced.
        builder.Writeln("No occurrence here.");                            // No match.

        // Save the document so that we can demonstrate loading it later.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Load the document and perform a conditional replace.
        // -----------------------------------------------------------------
        Document loaded = new Document(inputPath);

        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new StartOfParagraphReplacer()
        };

        // Replace the word "foo" with "bar" only when it is at the start of a paragraph.
        int replacedCount = loaded.Range.Replace("foo", "bar", options);

        // Validate that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement, but none were made.");

        // Save the modified document.
        loaded.Save(outputPath);
    }

    // -----------------------------------------------------------------
    // Callback that replaces a match only if it occurs at the start of a paragraph.
    // -----------------------------------------------------------------
    private class StartOfParagraphReplacer : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is at the start of its containing node when the offset is zero.
            // Additionally, ensure the match resides within a paragraph.
            if (args.MatchOffset == 0 && args.MatchNode?.ParentNode is Paragraph)
            {
                args.Replacement = "bar";
                return ReplaceAction.Replace;
            }

            // Skip any matches that are not at the beginning of a paragraph.
            return ReplaceAction.Skip;
        }
    }
}
