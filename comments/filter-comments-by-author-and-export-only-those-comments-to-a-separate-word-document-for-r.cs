using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class FilterCommentsExample
{
    public static void Main()
    {
        // Create a source document with several comments.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment by Alice.
        srcBuilder.Writeln("First paragraph with a comment.");
        Comment commentAlice = new Comment(sourceDoc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        commentAlice.AppendChild(new Paragraph(sourceDoc));
        commentAlice.FirstParagraph?.AppendChild(new Run(sourceDoc, "Please review this paragraph."));
        // Attach the comment to the current paragraph.
        srcBuilder.CurrentParagraph?.AppendChild(commentAlice);

        // Second paragraph with a comment by Bob.
        srcBuilder.Writeln("Second paragraph with a comment.");
        Comment commentBob = new Comment(sourceDoc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddMinutes(-10)
        };
        commentBob.AppendChild(new Paragraph(sourceDoc));
        commentBob.FirstParagraph?.AppendChild(new Run(sourceDoc, "Check the data accuracy."));
        srcBuilder.CurrentParagraph?.AppendChild(commentBob);

        // Third paragraph with another comment by Alice.
        srcBuilder.Writeln("Third paragraph with another comment.");
        Comment commentAlice2 = new Comment(sourceDoc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddHours(-1)
        };
        commentAlice2.AppendChild(new Paragraph(sourceDoc));
        commentAlice2.FirstParagraph?.AppendChild(new Run(sourceDoc, "Add more examples here."));
        srcBuilder.CurrentParagraph?.AppendChild(commentAlice2);

        // Save the source document (optional, for inspection).
        sourceDoc.Save("source.docx");

        // Filter comments authored by "Alice".
        var filteredComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, "Alice", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a new document to export the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln("Comments authored by Alice:");
        reportBuilder.Writeln();

        foreach (Comment c in filteredComments)
        {
            // Get the comment text; GetText may return null, so guard against it.
            string text = c.GetText()?.Trim() ?? string.Empty;

            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Date: {c.DateTime:O}");
            reportBuilder.Writeln($"Text: {text}");
            reportBuilder.Writeln();
        }

        // Save the report document.
        reportDoc.Save("alice-comments-report.docx");
    }
}
