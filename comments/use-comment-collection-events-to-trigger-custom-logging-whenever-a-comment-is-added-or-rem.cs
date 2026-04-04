using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class CommentEventDemo
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will hold comments.
        builder.Writeln("First paragraph – will receive a comment.");
        Paragraph firstParagraph = builder.CurrentParagraph;

        // Add a second paragraph for later removal demonstration.
        builder.Writeln("Second paragraph – will receive a comment that we later remove.");
        Paragraph secondParagraph = builder.CurrentParagraph;

        // Add comments using the helper method – this logs the addition.
        Comment comment1 = AddComment(doc, firstParagraph, "Alice", "A", "Review this sentence.");
        Comment comment2 = AddComment(doc, secondParagraph, "Bob", "B", "Consider rephrasing.");

        // Show current comments.
        Console.WriteLine("\nCurrent comments after additions:");
        ListComments(doc);

        // Remove the second comment using the helper method – this logs the removal.
        RemoveComment(doc, comment2);

        // Show remaining comments.
        Console.WriteLine("\nCurrent comments after removal:");
        ListComments(doc);

        // Save the document so the changes can be inspected manually if needed.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentEventDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"\nDocument saved to: {outputPath}");
    }

    // Creates a comment, attaches it to the specified paragraph, and logs the action.
    private static Comment AddComment(Document doc, Paragraph paragraph, string author, string initial, string text)
    {
        // Initialize a new comment with metadata.
        Comment comment = new Comment(doc, author, initial, DateTime.Now);
        comment.SetText(text);

        // Append the comment to the paragraph.
        paragraph.AppendChild(comment);

        // Log the addition.
        Console.WriteLine($"[Log] Added comment by '{author}': \"{text}\"");

        return comment;
    }

    // Removes a comment from the document and logs the action.
    private static void RemoveComment(Document doc, Comment comment)
    {
        if (comment == null)
        {
            Console.WriteLine("[Log] Attempted to remove a null comment – operation skipped.");
            return;
        }

        // Capture details before removal for logging.
        string author = comment.Author ?? "<unknown>";
        string text = comment.GetText()?.Trim() ?? "<empty>";

        // Perform removal.
        comment.Remove();

        // Log the removal.
        Console.WriteLine($"[Log] Removed comment by '{author}': \"{text}\"");
    }

    // Enumerates all comments in the document and writes their details to the console.
    private static void ListComments(Document doc)
    {
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        if (!comments.Any())
        {
            Console.WriteLine("  (No comments present)");
            return;
        }

        foreach (Comment c in comments)
        {
            string author = c.Author ?? "<unknown>";
            string text = c.GetText()?.Trim() ?? "<empty>";
            Console.WriteLine($"  Author: {author}, Text: \"{text}\"");
        }
    }
}
