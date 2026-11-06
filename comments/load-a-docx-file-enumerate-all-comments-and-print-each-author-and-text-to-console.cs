using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a temporary folder for the sample document.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Temp");
        Directory.CreateDirectory(tempFolder);

        // Path of the sample DOCX file.
        string samplePath = Path.Combine(tempFolder, "sample.docx");

        // Generate a DOCX file that contains at least one comment.
        CreateSampleDocument(samplePath);

        // Load the document from the file system.
        Document doc = new Document(samplePath);

        // Retrieve all comment nodes in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Print each comment's author and text to the console.
        foreach (Comment comment in comments)
        {
            // Trim whitespace from the comment text.
            string text = comment.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{comment.Author}: {text}");
        }
    }

    private static void CreateSampleDocument(string filePath)
    {
        // Initialize a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will host a comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create a comment node with metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add a paragraph and run inside the comment to hold its visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Append the comment to the last paragraph of the document.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Save the document to the specified location.
        doc.Save(filePath);
    }
}
