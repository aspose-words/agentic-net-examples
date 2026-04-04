using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will contain the comment anchor.
        builder.Writeln("This paragraph will have a comment with formatted text.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Build the comment body: a paragraph with bold, italic and normal runs.
        Paragraph commentParagraph = new Paragraph(doc);
        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        commentParagraph.AppendChild(boldRun);

        Run italicRun = new Run(doc, " Italic text");
        italicRun.Font.Italic = true;
        commentParagraph.AppendChild(italicRun);

        Run normalRun = new Run(doc, " Normal text.");
        commentParagraph.AppendChild(normalRun);

        // Add the paragraph to the comment story.
        comment.AppendChild(commentParagraph);

        // Anchor the comment to the current paragraph.
        Paragraph anchorParagraph = builder.CurrentParagraph;
        anchorParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        anchorParagraph.AppendChild(new Run(doc, "Commented text"));
        anchorParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        anchorParagraph.AppendChild(comment);

        // Insert a copy of the comment content into the main document body
        // so that it will be visible in the HTML output with its formatting preserved.
        Paragraph copiedComment = (Paragraph)commentParagraph.Clone(true);
        // Optionally add a label to indicate this is a comment.
        copiedComment.InsertBefore(new Run(doc, "Comment: "), copiedComment.FirstChild);
        // Insert after the paragraph that contains the comment anchor.
        anchorParagraph.ParentNode.InsertAfter(copiedComment, anchorParagraph);

        // Save the document in DOCX format (optional, for verification).
        doc.Save("output.docx");

        // Save the document as HTML. No special options are required.
        doc.Save("output.html", SaveFormat.Html);
    }
}
