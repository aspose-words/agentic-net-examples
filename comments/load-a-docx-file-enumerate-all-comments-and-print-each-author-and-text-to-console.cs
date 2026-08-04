using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a couple of comments.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // First paragraph.
        builder.Writeln("This is the first paragraph.");
        // Add a comment to the first paragraph.
        Comment comment1 = new Comment(sampleDoc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review this paragraph.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph.
        builder.Writeln("This is the second paragraph.");
        // Add a second comment.
        Comment comment2 = new Comment(sampleDoc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing this sentence.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the sample document to a temporary file.
        string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        sampleDoc.Save(tempFilePath);

        // Load the document from the file.
        Document loadedDoc = new Document(tempFilePath);

        // Enumerate all comments in the document.
        var comments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Print each comment's author and text.
        foreach (Comment c in comments)
        {
            string author = c.Author ?? string.Empty;
            string text = c.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{author}: {text}");
        }

        // Clean up the temporary file.
        if (File.Exists(tempFilePath))
        {
            File.Delete(tempFilePath);
        }
    }
}
