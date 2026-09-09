using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some paragraphs with comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with an old comment (2 months ago).
        builder.Writeln("First paragraph.");
        Comment oldComment = new Comment(doc, "Alice", "A", DateTime.Now.AddMonths(-2));
        oldComment.SetText("This is an old comment.");
        builder.CurrentParagraph.AppendChild(oldComment);

        // Second paragraph with a recent comment (today).
        builder.Writeln("Second paragraph.");
        Comment recentComment = new Comment(doc, "Bob", "B", DateTime.Now);
        recentComment.SetText("This is a recent comment.");
        builder.CurrentParagraph.AppendChild(recentComment);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the document before removal for reference.
        string beforePath = Path.Combine(outputDir, "CommentsBefore.docx");
        doc.Save(beforePath);

        // Define the date threshold: comments older than this date will be removed.
        DateTime threshold = DateTime.Now.AddMonths(-1); // 1 month ago

        // Enumerate all comment nodes safely and collect those older than the threshold.
        var oldComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .Where(c => c.DateTime < threshold)
                             .ToList();

        // Remove each old comment from the document.
        foreach (Comment comment in oldComments)
        {
            comment.Remove();
        }

        // Save the document after removal.
        string afterPath = Path.Combine(outputDir, "CommentsAfter.docx");
        doc.Save(afterPath);
    }
}
