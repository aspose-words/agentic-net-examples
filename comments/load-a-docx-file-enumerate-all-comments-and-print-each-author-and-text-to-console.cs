using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample document with a comment.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add a paragraph that will contain the comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Create the comment node.
        Comment comment = new Comment(sampleDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        comment.AppendChild(new Paragraph(sampleDoc));
        comment.FirstParagraph?.AppendChild(new Run(sampleDoc, "Please review this paragraph."));

        // Attach the comment to the first paragraph of the document.
        Paragraph? firstParagraph = sampleDoc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            firstParagraph.AppendChild(comment);
        }

        // Save the sample document to a local file.
        const string samplePath = "sample.docx";
        sampleDoc.Save(samplePath);

        // Load the document from the file.
        Document loadedDoc = new Document(samplePath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        // Print each comment's author and text to the console.
        foreach (Comment c in comments)
        {
            string text = c.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{c.Author}: {text}");
        }
    }
}
