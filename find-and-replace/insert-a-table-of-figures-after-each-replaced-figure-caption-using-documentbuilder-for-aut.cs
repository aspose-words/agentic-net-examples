using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Replacing;
using Aspose.Words.Tables;

public class InsertTableOfFiguresExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample figure captions and some regular text.
        builder.Writeln("Figure 1: First sample figure.");
        builder.Writeln("This is some introductory text.");
        builder.Writeln("Figure 2: Second sample figure.");
        builder.Writeln("More content follows.");
        builder.Writeln("Figure 3: Third sample figure.");

        // Save the original document (optional, just for reference).
        doc.Save("Original.docx");

        // Set up find-and-replace options with a custom callback.
        FindReplaceOptions options = new FindReplaceOptions
        {
            ReplacingCallback = new CaptionReplaceCallback()
        };

        // Replace the word "Figure" with "Fig." and trigger the callback for each match.
        int replacedCount = doc.Range.Replace("Figure", "Fig.", options);

        // Ensure that at least one replacement occurred.
        if (replacedCount == 0)
            throw new InvalidOperationException("No figure captions were replaced.");

        // Update fields (e.g., the inserted Table of Figures) before saving.
        doc.UpdateFields();

        // Save the modified document.
        doc.Save("Modified.docx");
    }

    // Callback that inserts a Table of Figures after each replaced caption.
    private class CaptionReplaceCallback : IReplacingCallback
    {
        public ReplaceAction Replacing(ReplacingArgs args)
        {
            // The match is inside a Run node; its parent is the Paragraph containing the caption.
            Paragraph captionParagraph = args.MatchNode.ParentNode as Paragraph;
            if (captionParagraph == null)
                return ReplaceAction.Skip;

            // Create a builder attached to the same document.
            DocumentBuilder cb = new DocumentBuilder((Document)args.MatchNode.Document);

            // Insert a new empty paragraph after the caption paragraph.
            Paragraph tocParagraph = new Paragraph(cb.Document);
            captionParagraph.ParentNode.InsertAfter(tocParagraph, captionParagraph);

            // Move the builder to the new paragraph and insert the Table of Figures field.
            cb.MoveTo(tocParagraph);
            // The field code "\c \"Caption\" \h \z \u" creates a Table of Figures for entries with the style "Caption".
            cb.InsertTableOfContents("\\c \"Caption\" \\h \\z \\u");

            // Continue with the normal replacement of the matched text.
            return ReplaceAction.Replace;
        }
    }
}
