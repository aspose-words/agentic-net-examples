using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file with a comment.
        string samplePath = Path.Combine(Path.GetTempPath(), "sample.docx");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph with a comment.");

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "This is a comment."));

        // Attach the comment to the paragraph we just wrote.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph != null)
        {
            paragraph.AppendChild(comment);
        }

        // Save the sample document.
        doc.Save(samplePath);

        // Load the DOCX file.
        Document loadedDoc = new Document(samplePath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Print each comment's author and text to the console.
        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author}: {c.GetText().Trim()}");
        }

        // Clean up the temporary file (optional).
        if (File.Exists(samplePath))
        {
            File.Delete(samplePath);
        }
    }
}
