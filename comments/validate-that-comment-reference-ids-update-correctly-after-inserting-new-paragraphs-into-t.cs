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

        // Write an initial paragraph that will contain the comment.
        builder.Writeln("Paragraph with comment.");

        // Retrieve the paragraph we just created.
        Paragraph? paragraph = doc.FirstSection?.Body?.FirstParagraph;
        if (paragraph == null)
        {
            Console.WriteLine("Failed to create the initial paragraph.");
            return;
        }

        // Create a top‑level comment.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a test comment.");

        // Anchor the comment to a range of text inside the paragraph.
        // The range consists of a CommentRangeStart, a Run, and a CommentRangeEnd.
        paragraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        paragraph.AppendChild(new Run(doc, "Commented text."));
        paragraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        paragraph.AppendChild(comment);

        // Store the original comment identifier.
        int originalId = comment.Id;

        // Insert new paragraphs before the paragraph that holds the comment.
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted paragraph 1.");
        builder.Writeln("Inserted paragraph 2.");

        // After insertion, enumerate all comments in the document.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        if (comments.Count == 0)
        {
            Console.WriteLine("No comments found after insertion.");
            return;
        }

        // There should be exactly one comment.
        Comment retrievedComment = comments[0];

        // Verify that the comment identifier has not changed.
        bool idUnchanged = retrievedComment.Id == originalId;

        // Verify that the surrounding CommentRangeStart and CommentRangeEnd share the same Id.
        // The expected order is: CommentRangeStart, Run, CommentRangeEnd, Comment.
        Node? rangeEndNode = retrievedComment.PreviousSibling;
        Node? rangeStartNode = rangeEndNode?.PreviousSibling?.PreviousSibling;

        bool rangeIdsMatch = false;
        if (rangeStartNode is CommentRangeStart start && rangeEndNode is CommentRangeEnd end)
        {
            rangeIdsMatch = start.Id == retrievedComment.Id && end.Id == retrievedComment.Id;
        }

        // Output verification results.
        Console.WriteLine($"Original Comment Id: {originalId}");
        Console.WriteLine($"Retrieved Comment Id: {retrievedComment.Id}");
        Console.WriteLine($"Comment Id unchanged after insertion: {idUnchanged}");
        Console.WriteLine($"CommentRangeStart and CommentRangeEnd IDs match comment Id: {rangeIdsMatch}");

        // Save the document to verify the result manually if needed.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentIdUpdate.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
