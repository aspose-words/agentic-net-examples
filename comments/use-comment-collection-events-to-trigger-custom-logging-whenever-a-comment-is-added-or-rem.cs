using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class CommentLogger
{
    // Adds a comment to the current paragraph of the builder and logs the action.
    public static Comment AddComment(Document doc, DocumentBuilder builder, string author, string initials, string text)
    {
        // Create the comment node.
        Comment comment = new Comment(doc, author, initials, DateTime.Now);
        comment.SetText(text);

        // Append the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(comment);

        // Log the addition.
        Console.WriteLine($"[Log] Added comment Id={comment.Id}, Author={author}, Text=\"{text}\"");
        return comment;
    }

    // Removes the specified comment from the document and logs the action.
    public static void RemoveComment(Comment comment)
    {
        if (comment == null)
        {
            Console.WriteLine("[Log] Attempted to remove a null comment – operation skipped.");
            return;
        }

        int id = comment.Id;
        string author = comment.Author ?? "<null>";
        string text = comment.GetText().Trim();

        // Remove the comment node.
        comment.Remove();

        // Log the removal.
        Console.WriteLine($"[Log] Removed comment Id={id}, Author={author}, Text=\"{text}\"");
    }
}

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will host comments.
        builder.Writeln("This is a sample paragraph for comment events demonstration.");

        // Add two comments using the logger.
        Comment firstComment = CommentLogger.AddComment(doc, builder, "Alice", "A", "First comment.");
        Comment secondComment = CommentLogger.AddComment(doc, builder, "Bob", "B", "Second comment.");

        // Enumerate current comments and display their details.
        Console.WriteLine("\n[Info] Current comments in the document:");
        var currentComments = doc.GetChildNodes(NodeType.Comment, true)
                                 .OfType<Comment>()
                                 .ToList();

        foreach (Comment c in currentComments)
        {
            Console.WriteLine($"- Id={c.Id}, Author={c.Author}, Text=\"{c.GetText().Trim()}\"");
        }

        // Remove the first comment using the logger.
        CommentLogger.RemoveComment(firstComment);

        // Enumerate comments after removal.
        Console.WriteLine("\n[Info] Comments after removal:");
        var remainingComments = doc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        foreach (Comment c in remainingComments)
        {
            Console.WriteLine($"- Id={c.Id}, Author={c.Author}, Text=\"{c.GetText().Trim()}\"");
        }

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CommentEventsDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"\nDocument saved to: {outputPath}");
    }
}
