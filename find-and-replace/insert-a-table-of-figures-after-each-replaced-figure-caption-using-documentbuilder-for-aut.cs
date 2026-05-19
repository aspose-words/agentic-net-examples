using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;

public class InsertTableOfFiguresAfterReplace
{
    public static void Main()
    {
        // Create a sample document with figure captions.
        const string inputPath = "input.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Introduction paragraph.");
        builder.Writeln("Figure 1: Sample");
        builder.Writeln("Some text between figures.");
        builder.Writeln("Figure 2: Sample");
        doc.Save(inputPath);

        // Load the document for processing.
        Document loaded = new Document(inputPath);

        // Set up a replace operation that replaces the word "Sample" with "Updated"
        // and inserts a Table of Figures after each caption.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new InsertTableAfterCaptionCallback()
        };

        int replacedCount = loaded.Range.Replace("Sample", "Updated", options);
        if (replacedCount == 0)
            throw new InvalidOperationException("Expected at least one replacement.");

        // Update fields so the inserted TOC (Table of Figures) shows its entries.
        loaded.UpdateFields();

        // Save the resulting document.
        const string outputPath = "output.docx";
        loaded.Save(outputPath);
    }

    // Callback that inserts a Table of Figures after the paragraph containing the matched caption.
    private class InsertTableAfterCaptionCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is part of a Run; its parent paragraph holds the figure caption.
            Paragraph? captionParagraph = args.MatchNode.ParentNode as Paragraph;
            if (captionParagraph == null)
                return ReplaceAction.Skip;

            // Use a DocumentBuilder to insert content after the caption paragraph.
            DocumentBuilder builder = new DocumentBuilder(args.MatchNode.Document as Document);

            // Move the cursor to the caption paragraph.
            builder.MoveTo(captionParagraph);
            // Insert a new empty paragraph after the caption.
            builder.InsertParagraph();

            // Insert a Table of Figures field that lists entries with the "Figure" label.
            builder.InsertField("TOC \\h \\z \\c \"Figure\"");

            // Continue with the normal replacement of the matched text.
            return ReplaceAction.Replace;
        }
    }
}
