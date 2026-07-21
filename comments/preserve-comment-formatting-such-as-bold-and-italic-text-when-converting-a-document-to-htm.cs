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

        // Add a sample paragraph that will host the comment.
        builder.Writeln("This is a sample paragraph.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Ensure the comment contains a paragraph.
        Paragraph commentParagraph = new Paragraph(doc);
        comment.AppendChild(commentParagraph);

        // Add a bold run inside the comment.
        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        commentParagraph.AppendChild(boldRun);

        // Add a space between runs.
        commentParagraph.AppendChild(new Run(doc, " "));

        // Add an italic run inside the comment.
        Run italicRun = new Run(doc, "Italic text");
        italicRun.Font.Italic = true;
        commentParagraph.AppendChild(italicRun);

        // Anchor the comment to the previously created paragraph.
        Paragraph? anchorParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (anchorParagraph != null)
        {
            anchorParagraph.AppendChild(comment);
        }

        // Export the comment's content to HTML while preserving formatting.
        string commentHtml = comment.ToString(SaveFormat.Html);
        File.WriteAllText("Comment.html", commentHtml);

        // Optionally, save the whole document for reference.
        doc.Save("DocumentWithComment.docx");
    }
}
