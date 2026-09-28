using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class CommentRemovalExample
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with an old comment (10 days ago).
        builder.Writeln("Paragraph with an old comment.");
        Paragraph paraOld = doc.FirstSection.Body.LastParagraph;
        Comment oldComment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddDays(-10) // older than threshold
        };
        // Add visible text to the comment.
        oldComment.AppendChild(new Paragraph(doc));
        oldComment.FirstParagraph?.AppendChild(new Run(doc, "This comment is older than the threshold."));
        // Attach the comment to the paragraph.
        paraOld?.AppendChild(oldComment);

        // Second paragraph with a recent comment (1 day ago).
        builder.Writeln("Paragraph with a recent comment.");
        Paragraph paraRecent = doc.FirstSection.Body.LastParagraph;
        Comment recentComment = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddDays(-1) // newer than threshold
        };
        recentComment.AppendChild(new Paragraph(doc));
        recentComment.FirstParagraph?.AppendChild(new Run(doc, "This comment is recent."));
        paraRecent?.AppendChild(recentComment);

        // Save the original document.
        string originalPath = "original.docx";
        doc.Save(originalPath);

        // Define the date threshold: comments older than this will be removed.
        DateTime threshold = DateTime.Now.AddDays(-5);

        // Collect all comments in a safe list.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        // Remove comments older than the threshold.
        foreach (Comment comment in allComments)
        {
            if (comment.DateTime < threshold)
            {
                comment.Remove();
            }
        }

        // Save the filtered document.
        string filteredPath = "filtered.docx";
        doc.Save(filteredPath);
    }
}
