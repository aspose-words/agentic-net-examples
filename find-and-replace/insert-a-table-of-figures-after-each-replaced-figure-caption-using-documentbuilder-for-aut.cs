using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample figure captions.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"Figure {i}: Sample caption for figure {i}.");
            // Add an empty paragraph to separate the figures.
            builder.Writeln();
        }

        // Set up find‑replace options with a custom callback that inserts a Table of Figures.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertTableOfFiguresHandler()
        };

        // Replace the word "Sample" with "Updated" in each caption.
        int replaceCount = doc.Range.Replace(new Regex("Sample"), "Updated", options);

        // Ensure that at least one replacement was performed.
        if (replaceCount == 0)
            throw new InvalidOperationException("No occurrences of the search text were found.");

        // Populate the inserted Table of Figures fields.
        doc.UpdateFields();

        // Save the resulting document.
        const string outputPath = "Result.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }

    // Callback that inserts a Table of Figures after each matched caption.
    private class InsertTableOfFiguresHandler : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // Find the paragraph that contains the match.
            Paragraph paragraph = args.MatchNode.GetAncestor(NodeType.Paragraph) as Paragraph;
            if (paragraph == null)
                return ReplaceAction.Skip;

            // Obtain the Document instance from the node (cast required for some Aspose.Words versions).
            Document doc = (Document)args.MatchNode.Document;

            // Use a builder positioned at the matched paragraph.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(paragraph);
            // Move the cursor after the paragraph.
            builder.Writeln();

            // Insert a Table of Figures (TOC with the \\f switch for "Figure" captions).
            builder.InsertTableOfContents("\\f \"Figure\" \\h \\z \\u");

            // Keep the original matched text unchanged.
            args.Replacement = args.Match.Value;
            return ReplaceAction.Replace;
        }
    }
}
