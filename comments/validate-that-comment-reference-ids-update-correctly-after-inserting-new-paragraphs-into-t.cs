using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing; // Required for CommentRangeStart/End

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add an initial paragraph that will contain a comment.
        builder.Writeln("Paragraph before comment.");

        // Create a top‑level comment with author metadata.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Now);
        comment.SetText("This is a comment.");

        // The comment must be anchored to a range of text.
        // Retrieve the paragraph that was just created by the builder.
        Paragraph? para = doc.FirstSection?.Body?.LastParagraph;

        if (para != null)
        {
            // Insert the comment range start, the commented text, the range end, and finally the comment node itself.
            para.AppendChild(new CommentRangeStart(doc, comment.Id));
            para.AppendChild(new Run(doc, "Commented text."));
            para.AppendChild(new CommentRangeEnd(doc, comment.Id));
            para.AppendChild(comment);
        }

        // Add another paragraph after the comment to have more content.
        builder.Writeln("Paragraph after comment.");

        // Save the initial document.
        doc.Save("CommentReferenceIds.docx");

        // Insert a new paragraph at the very beginning of the document.
        builder.MoveToDocumentStart();
        builder.Writeln("Inserted new paragraph at the beginning.");

        // Save the document after insertion.
        doc.Save("CommentReferenceIdsUpdated.docx");

        // -----------------------------------------------------------------
        // Validation: ensure that the comment's Id matches the Ids of its
        // associated CommentRangeStart and CommentRangeEnd nodes after the
        // insertion of new content.
        // -----------------------------------------------------------------
        var commentNodes = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

        if (commentNodes.Count == 0)
        {
            Console.WriteLine("No comments were found in the document.");
            return;
        }

        // There is only one comment in this example.
        Comment firstComment = commentNodes[0];
        int commentId = firstComment.Id;

        // The layout of nodes for a comment is:
        // CommentRangeStart, Run (commented text), CommentRangeEnd, Comment.
        // Therefore, the previous sibling of the comment is the CommentRangeEnd.
        CommentRangeEnd? rangeEnd = firstComment.PreviousSibling as CommentRangeEnd;

        // The CommentRangeStart is three nodes before the comment (skip Run and CommentRangeEnd).
        CommentRangeStart? rangeStart = null;
        if (rangeEnd != null && rangeEnd.PreviousSibling?.PreviousSibling != null)
        {
            rangeStart = rangeEnd.PreviousSibling.PreviousSibling as CommentRangeStart;
        }

        bool idsMatch = rangeStart != null && rangeEnd != null &&
                        rangeStart.Id == commentId && rangeEnd.Id == commentId;

        Console.WriteLine($"Comment Id: {commentId}");
        Console.WriteLine($"Range Start Id: {(rangeStart != null ? rangeStart.Id.ToString() : "null")}");
        Console.WriteLine($"Range End Id: {(rangeEnd != null ? rangeEnd.Id.ToString() : "null")}");
        Console.WriteLine($"IDs match after insertion: {idsMatch}");
    }
}
