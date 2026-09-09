using System;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some text to the document.
        builder.Writeln("Paragraph before comment.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This comment will be visible in the PDF output.");

        // Insert the comment into the first paragraph.
        // The comment is anchored to a range of text using CommentRangeStart and CommentRangeEnd.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        firstParagraph.AppendChild(new Run(doc, "Commented text"));
        firstParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        firstParagraph.AppendChild(comment);

        // Configure the layout to render comments as PDF annotations.
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.ShowInAnnotations;
        doc.UpdatePageLayout(); // Rebuild layout after changing options.

        // Save the document as PDF; comments will appear as visible annotations.
        doc.Save("DocumentWithComments.pdf", SaveFormat.Pdf);
    }
}
