using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph that will contain the comment.
        builder.Writeln("Paragraph before comment.");

        // Create a comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a sample comment.");

        // Anchor the comment to a range of text inside the first paragraph.
        Paragraph para = doc.FirstSection.Body.FirstParagraph;
        para.AppendChild(new CommentRangeStart(doc, comment.Id));
        para.AppendChild(new Run(doc, "Commented text."));
        para.AppendChild(new CommentRangeEnd(doc, comment.Id));
        para.AppendChild(comment);

        // Insert a new paragraph before the existing paragraph.
        DocumentBuilder insertBuilder = new DocumentBuilder(doc);
        insertBuilder.MoveToDocumentStart();
        insertBuilder.Writeln("Inserted paragraph before the commented paragraph.");

        // Verify that each comment's Id matches its range start and end identifiers.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        bool allMatch = true;
        foreach (Comment c in comments)
        {
            // The comment's previous sibling should be the CommentRangeEnd.
            CommentRangeEnd? rangeEnd = c.PreviousSibling as CommentRangeEnd;
            // The CommentRangeStart is three nodes before the comment.
            CommentRangeStart? rangeStart = c.PreviousSibling?.PreviousSibling?.PreviousSibling as CommentRangeStart;

            if (rangeStart == null || rangeEnd == null ||
                rangeStart.Id != c.Id || rangeEnd.Id != c.Id)
            {
                allMatch = false;
                break;
            }
        }

        Console.WriteLine(allMatch
            ? "All comment reference IDs are consistent after insertion."
            : "Comment reference IDs mismatch detected.");

        // Save the document for visual inspection.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentIdsUpdated.docx");
        doc.Save(outputPath);
    }
}
