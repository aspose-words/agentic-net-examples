using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new blank document.
        Document doc = new Document();

        // Build document content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This paragraph will receive a comment.");

        // Create a comment node.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        Paragraph commentParagraph = new Paragraph(doc);
        commentParagraph.AppendChild(new Run(doc, "Please review this paragraph."));
        comment.AppendChild(commentParagraph);

        // Attach the comment to the first paragraph.
        Paragraph? targetParagraph = doc.FirstSection?.Body?.FirstParagraph;
        if (targetParagraph != null)
        {
            targetParagraph.AppendChild(comment);
            Console.WriteLine($"[Log] Comment added – Author: {comment.Author}, Text: \"{comment.GetText().Trim()}\"");
        }

        // Remove the comment to demonstrate the removal event.
        comment.Remove();
        Console.WriteLine($"[Log] Comment removed – Author: {comment.Author}, Text: \"{comment.GetText().Trim()}\"");

        // Save the document.
        string outputPath = Path.Combine(outputDir, "CommentEventsDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
