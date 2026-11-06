using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph.
        builder.Writeln("Paragraph 1.");

        // Create a comment and attach it to the first paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "AL",
            DateTime = DateTime.Now
        };
        // Add visible text to the comment.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "This is a comment."));

        // Append the comment node to the first paragraph.
        Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
        firstParagraph?.AppendChild(comment);

        // Add a second paragraph after the comment.
        builder.Writeln("Paragraph 2.");

        // Capture the comment's Id before inserting a new paragraph.
        Comment? commentBefore = doc.GetChildNodes(NodeType.Comment, true)
                                    .OfType<Comment>()
                                    .FirstOrDefault();
        int? idBefore = commentBefore?.Id;

        // Insert a new paragraph at the very beginning of the document.
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted paragraph before the original comment.");

        // Capture the comment's Id after the insertion.
        Comment? commentAfter = doc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .FirstOrDefault();
        int? idAfter = commentAfter?.Id;

        // Output the Ids to verify they are unchanged.
        Console.WriteLine($"Comment Id before insertion: {idBefore}");
        Console.WriteLine($"Comment Id after insertion:  {idAfter}");
        Console.WriteLine($"Ids are {(idBefore == idAfter ? "consistent" : "different")} after insertion.");

        // Save the document to verify the result manually if needed.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentReferenceUpdate.docx");
        doc.Save(outputPath);
    }
}
