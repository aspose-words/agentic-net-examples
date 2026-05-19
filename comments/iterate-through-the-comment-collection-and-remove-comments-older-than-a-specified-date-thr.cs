using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph and an old comment (40 days ago).
        builder.Writeln("Paragraph with an old comment.");
        DateTime oldDate = DateTime.Now.AddDays(-40);
        Comment oldComment = new Comment(doc, "Alice", "A", oldDate);
        oldComment.SetText("This comment is older than the threshold.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(oldComment);

        // Add another paragraph and a recent comment (10 days ago).
        builder.Writeln("Paragraph with a recent comment.");
        DateTime recentDate = DateTime.Now.AddDays(-10);
        Comment recentComment = new Comment(doc, "Bob", "B", recentDate);
        recentComment.SetText("This comment is newer than the threshold.");
        builder.CurrentParagraph.AppendChild(recentComment);

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "comments_original.docx");
        doc.Save(originalPath);

        // Define the date threshold: comments older than 30 days will be removed.
        DateTime threshold = DateTime.Now.AddDays(-30);

        // Enumerate all comment nodes safely and collect those that are older than the threshold.
        var commentsToRemove = doc.GetChildNodes(NodeType.Comment, true)
                                  .OfType<Comment>()
                                  .Where(c => c.DateTime < threshold)
                                  .ToList();

        // Remove each matched comment from the document.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the filtered document.
        string filteredPath = Path.Combine(Directory.GetCurrentDirectory(), "comments_filtered.docx");
        doc.Save(filteredPath);
    }
}
