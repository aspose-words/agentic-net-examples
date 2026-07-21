using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add three paragraphs, each with a comment having a different date.
        AddParagraphWithComment(builder, "First paragraph.", "Alice", "A", DateTime.Now.AddDays(-5));
        AddParagraphWithComment(builder, "Second paragraph.", "Bob", "B", DateTime.Now.AddDays(-2));
        AddParagraphWithComment(builder, "Third paragraph.", "Charlie", "C", DateTime.Now);

        // Define the threshold: comments older than 3 days will be removed.
        DateTime threshold = DateTime.Now.AddDays(-3);

        // Collect all comment nodes safely into a list.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Remove comments whose DateTime is earlier than the threshold.
        foreach (Comment comment in comments)
        {
            if (comment.DateTime < threshold)
                comment.Remove();
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "CommentsFiltered.docx");
        doc.Save(outputPath);
    }

    // Helper method to add a paragraph with an attached comment.
    private static void AddParagraphWithComment(DocumentBuilder builder, string paragraphText,
                                                string author, string initial, DateTime commentDate)
    {
        // Write the paragraph text.
        builder.Writeln(paragraphText);

        // Create a new comment.
        Comment comment = new Comment(builder.Document, author, initial, commentDate);
        comment.SetText($"Comment by {author} on {commentDate:d}");

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);
    }
}
