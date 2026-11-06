using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will contain a comment.
        builder.Writeln("This is a paragraph that will have a comment attached.");

        // Create a comment with custom author name and initials.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // The comment must contain at least one paragraph and run so that it has visible text.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the last paragraph that was added by the builder.
        Paragraph? targetParagraph = builder.CurrentParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
        }

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);

        // Reload the document to demonstrate that the comment metadata was saved correctly.
        Document loadedDoc = new Document(outputPath);

        // Enumerate all comments in the document.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            // Output author, initials, and comment text to the console.
            Console.WriteLine($"Author: {c.Author}");
            Console.WriteLine($"Initials: {c.Initial}");
            Console.WriteLine($"Date: {c.DateTime}");
            Console.WriteLine($"Text: {c.GetText().Trim()}");
            Console.WriteLine();
        }
    }
}
