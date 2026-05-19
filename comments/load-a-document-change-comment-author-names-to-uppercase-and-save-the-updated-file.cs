using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "input.docx");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");

        // Create a sample document that contains a comment.
        CreateSampleDocument(inputPath);

        // Load the document from the file system.
        Document doc = new Document(inputPath);

        // Enumerate all comment nodes safely.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Convert each comment author name to uppercase.
        foreach (Comment comment in comments)
        {
            if (!string.IsNullOrEmpty(comment.Author))
                comment.Author = comment.Author.ToUpperInvariant();
        }

        // Save the modified document.
        doc.Save(outputPath);
    }

    private static void CreateSampleDocument(string filePath)
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with some text.
        builder.Writeln("This is a sample paragraph with a comment.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Review this paragraph.");

        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document to the specified path.
        doc.Save(filePath);
    }
}
