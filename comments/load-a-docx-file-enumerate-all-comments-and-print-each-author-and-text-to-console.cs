using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare a sample DOCX file with comments.
        string fileName = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        CreateSampleDocumentWithComments(fileName);

        // Load the DOCX file.
        Document doc = new Document(fileName);

        // Enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        // Print author and text of each comment.
        foreach (Comment comment in comments)
        {
            string author = comment.Author ?? "Unknown";
            string text = comment.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{author}: {text}");
        }
    }

    private static void CreateSampleDocumentWithComments(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph.
        builder.Writeln("This is the first paragraph.");

        // Add a comment.
        Comment comment1 = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(doc));
        comment1.FirstParagraph?.AppendChild(new Run(doc, "First comment text."));
        // Attach comment to the document body.
        doc.FirstSection?.Body?.FirstParagraph?.AppendChild(comment1);

        // Second paragraph.
        builder.Writeln("This is the second paragraph.");

        // Add another comment.
        Comment comment2 = new Comment(doc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(doc));
        comment2.FirstParagraph?.AppendChild(new Run(doc, "Second comment text."));
        doc.FirstSection?.Body?.LastParagraph?.AppendChild(comment2);

        // Save the document.
        doc.Save(path);
    }
}
