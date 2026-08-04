using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold the comments.
        builder.Writeln("Sample paragraph for comments.");

        // Create an old comment (10 days ago) – this should be removed.
        DateTime oldDate = DateTime.Now.AddDays(-10);
        Comment oldComment = new Comment(doc, "Old Author", "OA", oldDate);
        oldComment.SetText("This comment is older than the threshold.");
        builder.CurrentParagraph.AppendChild(oldComment);

        // Create a recent comment (today) – this should be kept.
        DateTime recentDate = DateTime.Now;
        Comment recentComment = new Comment(doc, "New Author", "NA", recentDate);
        recentComment.SetText("This comment is within the threshold.");
        builder.CurrentParagraph.AppendChild(recentComment);

        // Define the date threshold: comments older than this will be removed.
        DateTime threshold = DateTime.Now.AddDays(-5);

        // Enumerate all comment nodes safely (make a copy to avoid modifying the collection while iterating).
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        // Remove comments whose DateTime is earlier than the threshold.
        foreach (Comment comment in allComments)
        {
            if (comment.DateTime < threshold)
                comment.Remove();
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "comments_filtered.docx");
        doc.Save(outputPath);
    }
}
