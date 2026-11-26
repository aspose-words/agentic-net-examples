using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document with two paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Helper methods for logging comment actions.
        void LogCommentAdded(Comment c) =>
            Console.WriteLine($"[Log] Comment added by {c.Author} at {c.DateTime:u}");

        void LogCommentRemoved(Comment c) =>
            Console.WriteLine($"[Log] Comment removed by {c.Author} (original date {c.DateTime:u})");

        // Add a comment to the first paragraph.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph as Paragraph;
        if (firstParagraph != null)
        {
            Comment comment1 = new Comment(doc)
            {
                Author = "Alice",
                Initial = "A",
                DateTime = DateTime.Now
            };
            // A comment must contain at least one paragraph.
            comment1.AppendChild(new Paragraph(doc));
            // Add visible text to the comment.
            comment1.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));
            // Append the comment to the target paragraph.
            firstParagraph.AppendChild(comment1);
            LogCommentAdded(comment1);
        }

        // Add a comment to the second paragraph.
        Paragraph? secondParagraph = doc.FirstSection?.Body?.Paragraphs.ElementAtOrDefault(1) as Paragraph;
        if (secondParagraph != null)
        {
            Comment comment2 = new Comment(doc)
            {
                Author = "Bob",
                Initial = "B",
                DateTime = DateTime.Now.AddMinutes(-5)
            };
            comment2.AppendChild(new Paragraph(doc));
            comment2.FirstParagraph?.AppendChild(new Run(doc, "Check the data here."));
            secondParagraph.AppendChild(comment2);
            LogCommentAdded(comment2);
        }

        // Remove the first comment to trigger the removal log.
        Comment? commentToRemove = doc.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .FirstOrDefault(c => c.Author == "Alice");
        if (commentToRemove != null)
        {
            LogCommentRemoved(commentToRemove);
            commentToRemove.Remove();
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsEventsDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
