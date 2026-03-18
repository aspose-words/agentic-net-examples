using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add the first paragraph.
        builder.Writeln("First paragraph.");

        // Insert a comment anchored to the first paragraph.
        Comment comment = new Comment(doc, "Alice", "A", DateTime.Today);
        Paragraph firstPara = doc.FirstSection.Body.FirstParagraph;

        // Build the comment range: start -> run -> end -> comment node.
        firstPara.AppendChild(new CommentRangeStart(doc, comment.Id));
        firstPara.AppendChild(new Run(doc, "Commented text."));
        firstPara.AppendChild(new CommentRangeEnd(doc, comment.Id));
        firstPara.AppendChild(comment);
        comment.SetText("Initial comment.");

        // Capture the original IDs for later comparison.
        int originalCommentId = comment.Id;
        int originalStartId = ((CommentRangeStart)comment.PreviousSibling.PreviousSibling.PreviousSibling).Id;
        int originalEndId = ((CommentRangeEnd)comment.PreviousSibling).Id;

        // Insert a new paragraph before the existing one.
        builder.MoveTo(doc.FirstSection.Body.FirstParagraph);
        builder.InsertParagraph();               // Inserts an empty paragraph before the current.
        builder.Writeln("Inserted paragraph."); // Fill the new paragraph with text.

        // Retrieve the comment after insertion.
        Comment updatedComment = (Comment)doc.GetChildNodes(NodeType.Comment, true)[0];
        int updatedCommentId = updatedComment.Id;
        int updatedStartId = ((CommentRangeStart)updatedComment.PreviousSibling.PreviousSibling.PreviousSibling).Id;
        int updatedEndId = ((CommentRangeEnd)updatedComment.PreviousSibling).Id;

        // Validate that the comment ID matches its range start and end IDs.
        if (updatedCommentId != updatedStartId || updatedCommentId != updatedEndId)
            throw new InvalidOperationException("Comment range IDs are out of sync after inserting paragraphs.");

        // Output the IDs before and after insertion for verification.
        Console.WriteLine($"Original IDs: Comment={originalCommentId}, Start={originalStartId}, End={originalEndId}");
        Console.WriteLine($"After insertion: Comment={updatedCommentId}, Start={updatedStartId}, End={updatedEndId}");

        // Save the document.
        doc.Save("CommentReferenceUpdate.docx");
    }
}
