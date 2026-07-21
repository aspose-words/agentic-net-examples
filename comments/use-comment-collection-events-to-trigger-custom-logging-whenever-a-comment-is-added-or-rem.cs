using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    // Simple logger that writes messages to the console.
    private static void Log(string message)
    {
        Console.WriteLine($"[Log] {message}");
    }

    public static void Main()
    {
        // Ensure the Aspose.Words license is not required for this example.
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will contain a comment.
        builder.Writeln("This is a sample paragraph for comment demonstration.");

        // Create a comment and attach it to the paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Please review this paragraph.");
        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Log the addition of the comment.
        Log($"Comment added by '{comment.Author}' with text: \"{comment.GetText().Trim()}\"");

        // Save the document before removal to demonstrate both states.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string beforePath = Path.Combine(outputDir, "DocumentWithComment.docx");
        doc.Save(beforePath);
        Log($"Document saved with comment: {beforePath}");

        // Retrieve all comment nodes safely.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Remove each comment while logging the removal.
        foreach (Comment c in comments)
        {
            // Capture author before removal.
            string author = c.Author;
            string text = c.GetText().Trim();

            // Remove the comment from the document.
            c.Remove();

            // Log the removal.
            Log($"Comment removed by '{author}' with text: \"{text}\"");
        }

        // Save the document after removal.
        string afterPath = Path.Combine(outputDir, "DocumentWithoutComment.docx");
        doc.Save(afterPath);
        Log($"Document saved after comment removal: {afterPath}");
    }
}
