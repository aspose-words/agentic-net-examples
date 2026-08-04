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

        // Add the original paragraph that will contain the comment.
        builder.Writeln("First paragraph with comment.");

        // Create a comment anchored to the above paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("Initial comment.");

        // Build the comment range: start, commented text, end, then the comment itself.
        Paragraph? para = doc.FirstSection?.Body?.FirstParagraph;
        if (para == null)
        {
            Console.WriteLine("Failed to locate the first paragraph.");
            return;
        }

        para.AppendChild(new CommentRangeStart(doc, comment.Id));
        para.AppendChild(new Run(doc, "Commented text."));
        para.AppendChild(new CommentRangeEnd(doc, comment.Id));
        para.AppendChild(comment);

        // Insert a new paragraph before the paragraph that holds the comment.
        Paragraph newParagraph = new Paragraph(doc);
        newParagraph.AppendChild(new Run(doc, "Inserted paragraph before comment."));
        doc.FirstSection?.Body?.InsertBefore(newParagraph, para);

        // Verify that the comment ID matches its range start and end IDs.
        Comment? retrievedComment = doc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .FirstOrDefault();

        if (retrievedComment == null)
        {
            Console.WriteLine("No comment found in the document.");
            return;
        }

        int commentId = retrievedComment.Id;

        var rangeStart = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                            .OfType<CommentRangeStart>()
                            .FirstOrDefault(r => r.Id == commentId);

        var rangeEnd = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                          .OfType<CommentRangeEnd>()
                          .FirstOrDefault(r => r.Id == commentId);

        bool idsMatch = rangeStart != null && rangeEnd != null &&
                        rangeStart.Id == commentId && rangeEnd.Id == commentId;

        Console.WriteLine($"Comment ID: {commentId}");
        Console.WriteLine($"Range Start ID: {(rangeStart != null ? rangeStart.Id.ToString() : "null")}");
        Console.WriteLine($"Range End ID: {(rangeEnd != null ? rangeEnd.Id.ToString() : "null")}");
        Console.WriteLine($"IDs consistent: {idsMatch}");

        // Save the document to verify the result manually if needed.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "CommentIdUpdate.docx");
        doc.Save(outputPath);
    }
}
