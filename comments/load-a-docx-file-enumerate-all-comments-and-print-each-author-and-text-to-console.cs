using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary folder for the sample document.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Temp");
        Directory.CreateDirectory(tempFolder);

        // Path of the sample DOCX file.
        string samplePath = Path.Combine(tempFolder, "SampleWithComments.docx");

        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review this paragraph.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with another comment.
        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(doc, "Bob", "B", DateTime.Now);
        comment2.SetText("Consider rephrasing.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the document to disk.
        doc.Save(samplePath);

        // Load the document from the file system.
        Document loadedDoc = new Document(samplePath);

        // Enumerate all comments in the document.
        var comments = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        // Print author and comment text for each comment.
        foreach (Comment c in comments)
        {
            // GetText() returns the comment text including any trailing line breaks.
            string text = c.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{c.Author}: {text}");
        }
    }
}
