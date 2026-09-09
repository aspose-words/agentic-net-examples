using System;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add the first paragraph that will later contain a comment.
        builder.Writeln("This is the original paragraph that will be commented.");

        // Retrieve the paragraph we just added.
        Paragraph originalParagraph = doc.FirstSection.Body.FirstParagraph;

        // Create a comment and set its metadata.
        Comment comment = new Comment(doc)
        {
            Author = "Alice",
            Initial = "A",
            DateTime = DateTime.Now
        };
        // A comment must contain at least one paragraph with some text.
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph.AppendChild(new Run(doc, "Original comment text."));

        // Store the comment identifier before adding it to the document.
        int originalCommentId = comment.Id;

        // Anchor the comment to a range of text inside the original paragraph.
        originalParagraph.AppendChild(new CommentRangeStart(doc, comment.Id));
        originalParagraph.AppendChild(new Run(doc, "Commented text."));
        originalParagraph.AppendChild(new CommentRangeEnd(doc, comment.Id));
        originalParagraph.AppendChild(comment);

        // Save the document before modification (optional, for inspection).
        doc.Save("CommentBeforeInsertion.docx");

        // Insert a new paragraph before the original paragraph.
        Paragraph insertedParagraph = new Paragraph(doc);
        insertedParagraph.AppendChild(new Run(doc, "This is a newly inserted paragraph."));

        // Insert the new paragraph into the document tree.
        // The parent of a paragraph is a Body, which derives from CompositeNode and supports InsertBefore.
        CompositeNode? parent = originalParagraph.ParentNode as CompositeNode;
        if (parent != null)
        {
            parent.InsertBefore(insertedParagraph, originalParagraph);
        }

        // After insertion, retrieve the comment again.
        Comment? retrievedComment = doc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .FirstOrDefault();

        // Validate that the comment identifier has not changed.
        bool idUnchanged = retrievedComment != null && retrievedComment.Id == originalCommentId;

        // Verify that the comment range start and end nodes still reference the same identifier.
        bool rangeStartMatches = doc.GetChildNodes(NodeType.CommentRangeStart, true)
                                    .OfType<CommentRangeStart>()
                                    .Any(crs => crs.Id == originalCommentId);
        bool rangeEndMatches = doc.GetChildNodes(NodeType.CommentRangeEnd, true)
                                  .OfType<CommentRangeEnd>()
                                  .Any(cre => cre.Id == originalCommentId);

        // Output validation results.
        Console.WriteLine($"Original comment Id: {originalCommentId}");
        Console.WriteLine($"Comment Id after insertion: {(retrievedComment?.Id.ToString() ?? "null")}");
        Console.WriteLine($"Comment Id unchanged: {idUnchanged}");
        Console.WriteLine($"CommentRangeStart Id matches comment: {rangeStartMatches}");
        Console.WriteLine($"CommentRangeEnd Id matches comment: {rangeEndMatches}");

        // Save the final document.
        doc.Save("CommentAfterInsertion.docx");
    }
}
