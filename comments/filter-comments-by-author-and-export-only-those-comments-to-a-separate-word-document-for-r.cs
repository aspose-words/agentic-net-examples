using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a source document with several comments from different authors.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment from Alice.
        srcBuilder.Writeln("First paragraph with a comment.");
        Comment commentAlice = new Comment(sourceDoc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        commentAlice.AppendChild(new Paragraph(sourceDoc));
        commentAlice.FirstParagraph?.AppendChild(new Run(sourceDoc, "Alice's review note."));
        // Attach the comment to the first paragraph.
        sourceDoc.FirstSection?.Body?.FirstParagraph?.AppendChild(commentAlice);

        // Second paragraph with a comment from Bob.
        srcBuilder.Writeln("Second paragraph with a comment.");
        Comment commentBob = new Comment(sourceDoc)
        {
            Author = "Bob",
            Initial = "B",
            DateTime = DateTime.Now.AddMinutes(-10)
        };
        commentBob.AppendChild(new Paragraph(sourceDoc));
        commentBob.FirstParagraph?.AppendChild(new Run(sourceDoc, "Bob's suggestion."));
        sourceDoc.FirstSection?.Body?.LastParagraph?.AppendChild(commentBob);

        // Third paragraph with another comment from Alice.
        srcBuilder.Writeln("Third paragraph with another comment.");
        Comment commentAlice2 = new Comment(sourceDoc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now.AddHours(-1)
        };
        commentAlice2.AppendChild(new Paragraph(sourceDoc));
        commentAlice2.FirstParagraph?.AppendChild(new Run(sourceDoc, "Second note from Alice."));
        sourceDoc.FirstSection?.Body?.LastParagraph?.AppendChild(commentAlice2);

        // Save the source document (optional, for inspection).
        sourceDoc.Save("source.docx");

        // Define the author whose comments we want to export.
        const string targetAuthor = "Alice";

        // Enumerate all comments and filter by author.
        var filteredComments = sourceDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a new document to hold the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        reportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        reportBuilder.Writeln();

        foreach (Comment c in filteredComments)
        {
            reportBuilder.Writeln($"Author: {c.Author}");
            reportBuilder.Writeln($"Date: {c.DateTime:O}");
            reportBuilder.Writeln($"Text: {c.GetText().Trim()}");
            reportBuilder.Writeln(); // Add an empty line between comments.
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Save the report document.
        string reportPath = Path.Combine(outputDir, "filtered-comments.docx");
        reportDoc.Save(reportPath);
    }
}
