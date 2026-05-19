using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class PreserveCommentFormatting
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain the commented text.
        builder.Writeln("This paragraph will have a comment with formatted text.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);

        // Ensure the comment has a paragraph to hold its content.
        comment.AppendChild(new Paragraph(doc));

        // Add a bold run to the comment.
        Run boldRun = new Run(doc, "Bold");
        boldRun.Font.Bold = true;
        comment.FirstParagraph?.AppendChild(boldRun);

        // Add an italic run to the comment (preceded by a space for readability).
        Run italicRun = new Run(doc, " Italic");
        italicRun.Font.Italic = true;
        comment.FirstParagraph?.AppendChild(italicRun);

        // Anchor the comment to a range of text in the document.
        // Insert the start of the comment range.
        builder.CurrentParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        // The text that the comment refers to.
        builder.Write("Commented text");
        // Insert the end of the comment range.
        builder.CurrentParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        // Finally, add the comment node itself.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to HTML. Formatting inside the comment (bold/italic) is preserved automatically.
        string htmlPath = "CommentWithFormatting.html";
        doc.Save(htmlPath, SaveFormat.Html);

        // Enumerate all comments to demonstrate that they are present and contain the expected text.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // GetText returns the plain text of the comment (formatting is not shown in plain text).
            Console.WriteLine($"Comment by {c.Author}: {c.GetText().Trim()}");
        }

        Console.WriteLine($"HTML document saved to: {Path.GetFullPath(htmlPath)}");
    }
}
