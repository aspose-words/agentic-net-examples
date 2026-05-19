using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("This is the first paragraph of the document.");
        builder.Writeln("This is the second paragraph of the document.");

        // Add comments with custom logging.
        AddCommentWithLogging(doc, builder.CurrentParagraph, "Alice", "AL", "Review the first paragraph.");
        AddCommentWithLogging(doc, builder.CurrentParagraph, "Bob", "B", "Check the second paragraph for accuracy.");

        // List all comments currently in the document.
        Console.WriteLine("\nCurrent comments in the document:");
        foreach (Comment c in GetAllComments(doc))
        {
            Console.WriteLine($"- Author: {c.Author}, Text: \"{c.GetText().Trim()}\"");
        }

        // Remove the first comment (if any) with custom logging.
        Comment? firstComment = GetAllComments(doc).FirstOrDefault();
        if (firstComment != null)
        {
            RemoveCommentWithLogging(doc, firstComment);
        }

        // List comments after removal.
        Console.WriteLine("\nComments after removal:");
        foreach (Comment c in GetAllComments(doc))
        {
            Console.WriteLine($"- Author: {c.Author}, Text: \"{c.GetText().Trim()}\"");
        }

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CommentsDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"\nDocument saved to: {outputPath}");
    }

    // Retrieves all comment nodes in the document safely.
    private static System.Collections.Generic.List<Comment> GetAllComments(Document doc)
    {
        return doc.GetChildNodes(NodeType.Comment, true)
                  .OfType<Comment>()
                  .ToList();
    }

    // Adds a comment to the specified paragraph and logs the action.
    private static void AddCommentWithLogging(Document doc, Paragraph paragraph, string author, string initial, string text)
    {
        // Create a new comment with metadata.
        Comment comment = new Comment(doc, author, initial, DateTime.Now);
        comment.SetText(text);

        // Append the comment to the paragraph.
        paragraph.AppendChild(comment);

        // Log the addition.
        Console.WriteLine($"[LOG] Added comment by '{author}': \"{text}\"");
    }

    // Removes a comment from the document and logs the action.
    private static void RemoveCommentWithLogging(Document doc, Comment comment)
    {
        // Store details for logging before removal.
        string author = comment.Author ?? "Unknown";
        string text = comment.GetText().Trim();

        // Remove the comment safely.
        comment.Remove();

        // Log the removal.
        Console.WriteLine($"[LOG] Removed comment by '{author}': \"{text}\"");
    }
}
