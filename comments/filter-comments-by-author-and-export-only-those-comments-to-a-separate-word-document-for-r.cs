using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with several comments.
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // First paragraph with a comment by John Doe.
        srcBuilder.Writeln("First paragraph with a comment.");
        Comment comment1 = new Comment(sourceDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment1.AppendChild(new Paragraph(sourceDoc));
        comment1.FirstParagraph?.AppendChild(new Run(sourceDoc, "Please review this paragraph."));
        Paragraph? para1 = sourceDoc.FirstSection?.Body?.LastParagraph;
        para1?.AppendChild(comment1);

        // Second paragraph with a comment by Jane Smith.
        srcBuilder.Writeln("Second paragraph with a different comment.");
        Comment comment2 = new Comment(sourceDoc)
        {
            Author = "Jane Smith",
            Initial = "JS",
            DateTime = DateTime.Now.AddMinutes(-5)
        };
        comment2.AppendChild(new Paragraph(sourceDoc));
        comment2.FirstParagraph?.AppendChild(new Run(sourceDoc, "Consider rephrasing this sentence."));
        Paragraph? para2 = sourceDoc.FirstSection?.Body?.LastParagraph;
        para2?.AppendChild(comment2);

        // Third paragraph with another comment by John Doe.
        srcBuilder.Writeln("Third paragraph, another note.");
        Comment comment3 = new Comment(sourceDoc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now.AddHours(-1)
        };
        comment3.AppendChild(new Paragraph(sourceDoc));
        comment3.FirstParagraph?.AppendChild(new Run(sourceDoc, "Add a citation here."));
        Paragraph? para3 = sourceDoc.FirstSection?.Body?.LastParagraph;
        para3?.AppendChild(comment3);

        // Define the author to filter.
        const string targetAuthor = "John Doe";

        // Retrieve all comments from the source document.
        var allComments = sourceDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        // Filter comments by the specified author (case‑insensitive).
        var filteredComments = allComments
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a new document to hold the filtered comments.
        Document reportDoc = new Document();
        DocumentBuilder reportBuilder = new DocumentBuilder(reportDoc);

        // Write a header.
        reportBuilder.Writeln($"Comments authored by \"{targetAuthor}\":");
        reportBuilder.Writeln();

        // Add each filtered comment to the report document.
        foreach (Comment c in filteredComments)
        {
            string commentText = c.GetText().Trim();
            reportBuilder.Writeln($"Author : {c.Author}");
            reportBuilder.Writeln($"Date   : {c.DateTime:O}");
            reportBuilder.Writeln($"Text   : {commentText}");
            reportBuilder.Writeln(); // Blank line between comments.
        }

        // Ensure the output directory exists and save the report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FilteredComments.docx");
        reportDoc.Save(outputPath);
        Console.WriteLine($"Filtered comments saved to: {outputPath}");
    }
}
