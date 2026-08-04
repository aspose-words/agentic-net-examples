using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a sample paragraph that will have a comment.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Append the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(comment);

        // Build the comment body with formatted runs.
        Paragraph commentParagraph = comment.AppendChild(new Paragraph(doc));

        Run boldRun = new Run(doc, "Bold text");
        boldRun.Font.Bold = true;
        commentParagraph.AppendChild(boldRun);

        commentParagraph.AppendChild(new Run(doc, " and "));

        Run italicRun = new Run(doc, "italic text");
        italicRun.Font.Italic = true;
        commentParagraph.AppendChild(italicRun);

        // Save the document to HTML. The comment's formatting is preserved in the output.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentWithFormatting.html");
        doc.Save(outputPath, SaveFormat.Html);
    }
}
