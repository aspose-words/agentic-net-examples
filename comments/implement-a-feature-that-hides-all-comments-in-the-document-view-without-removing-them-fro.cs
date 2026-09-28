using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph that will have a comment.
        builder.Writeln("Paragraph with a comment.");

        // Create a comment node with author metadata and visible text.
        Comment comment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

        // Attach the comment to the first paragraph of the document.
        doc.FirstSection?.Body?.FirstParagraph?.AppendChild(comment);

        // Hide all comments in the document view (they remain in the file).
        doc.LayoutOptions.CommentDisplayMode = CommentDisplayMode.Hide;

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        doc.Save(outputPath);
    }
}
