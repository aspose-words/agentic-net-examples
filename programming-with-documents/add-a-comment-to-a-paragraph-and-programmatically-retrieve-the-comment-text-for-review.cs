using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add a paragraph with some text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph.");

        // Create a comment anchored to the current paragraph.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Review this paragraph for clarity.");

        // Append the comment to the paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CommentExample.docx");
        doc.Save(outputPath);

        // Retrieve all top‑level comments from the document.
        NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
        foreach (Comment c in commentNodes)
        {
            // Output the comment text and its author.
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Comment: {c.GetText().Trim()}");
        }
    }
}
