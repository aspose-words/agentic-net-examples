using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class CommentEventsDemo
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph that will host a comment.
        builder.Writeln("This is a paragraph that will have a comment.");

        // Get the paragraph we just added.
        Paragraph? paragraph = builder.CurrentParagraph;
        if (paragraph == null)
        {
            Console.WriteLine("Failed to obtain the paragraph.");
            return;
        }

        // Create a comment node with metadata.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };

        // Add a paragraph and run inside the comment so it contains visible text.
        comment.AppendChild(new Paragraph(doc));
        Paragraph? commentParagraph = comment.FirstParagraph;
        if (commentParagraph != null)
        {
            commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        }

        // Append the comment to the paragraph and log the addition.
        paragraph.AppendChild(comment);
        Console.WriteLine($"[Log] Comment added by '{comment.Author}'.");

        // Remove the comment and log the removal.
        comment.Remove();
        Console.WriteLine($"[Log] Comment removed (original author: '{comment.Author}').");

        // Save the document.
        doc.Save("CommentEventsDemo.docx");
    }
}
