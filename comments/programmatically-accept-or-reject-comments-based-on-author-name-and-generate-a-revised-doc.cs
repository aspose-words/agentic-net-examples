using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with several comments.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from John Doe.
        builder.Writeln("This is the first paragraph.");
        AddComment(builder, doc, "John Doe", "JD", "Review the first paragraph.");

        // Second paragraph with a comment from Jane Smith.
        builder.Writeln("This is the second paragraph.");
        AddComment(builder, doc, "Jane Smith", "JS", "Check the second paragraph.");

        // Third paragraph with a comment from Alice Brown.
        builder.Writeln("This is the third paragraph.");
        AddComment(builder, doc, "Alice Brown", "AB", "Consider revising the third paragraph.");

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Define the author whose comments we want to keep.
        const string targetAuthor = "John Doe";

        // Enumerate all comment nodes safely.
        var allComments = doc.GetChildNodes(NodeType.Comment, true)
                             .OfType<Comment>()
                             .ToList();

        // Remove comments that are not authored by the target author.
        var commentsToRemove = allComments
            .Where(c => !string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (Comment comment in commentsToRemove)
        {
            // Ensure the comment node is still attached before removal.
            if (comment.ParentNode != null)
                comment.Remove();
        }

        // Save the revised document containing only the kept comments.
        string revisedPath = Path.Combine(outputDir, "revised.docx");
        doc.Save(revisedPath);
    }

    // Helper method to create and attach a comment to the current paragraph.
    private static void AddComment(DocumentBuilder builder, Document doc, string author, string initials, string text)
    {
        // Create a new comment with metadata.
        Comment comment = new Comment(doc, author, initials, DateTime.Now);
        comment.SetText(text);

        // Append the comment to the current paragraph if it exists.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph != null)
        {
            paragraph.AppendChild(comment);
        }
    }
}
