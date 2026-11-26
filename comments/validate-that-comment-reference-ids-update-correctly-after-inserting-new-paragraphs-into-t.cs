using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and add initial content with a comment.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph that will receive the comment.
        builder.Writeln("Original paragraph with comment.");

        // Retrieve the paragraph we just created.
        Paragraph? originalParagraph = doc.FirstSection?.Body?.LastParagraph;
        if (originalParagraph == null)
        {
            Console.WriteLine("Failed to locate the original paragraph.");
            return;
        }

        // Create a comment and attach it to the original paragraph.
        Comment comment = new Comment(doc)
        {
            Author = "John Doe",
            Initial = "JD",
            DateTime = DateTime.Now
        };
        comment.AppendChild(new Paragraph(doc));
        comment.FirstParagraph?.AppendChild(new Run(doc, "This is a review comment."));
        originalParagraph.AppendChild(comment);

        // Another paragraph after the comment.
        builder.Writeln("Paragraph after the comment.");

        // Save the initial document.
        const string initialFile = "initial.docx";
        doc.Save(initialFile);

        // Retrieve the comment we just created.
        Comment? retrievedComment = doc.GetChildNodes(NodeType.Comment, true)
                                      .OfType<Comment>()
                                      .FirstOrDefault();

        if (retrievedComment == null)
        {
            Console.WriteLine("No comment found in the document.");
            return;
        }

        int originalCommentId = retrievedComment.Id;

        // Insert a new paragraph BEFORE the original paragraph.
        Paragraph newParagraph = new Paragraph(doc);
        newParagraph.AppendChild(new Run(doc, "Inserted new paragraph before the commented paragraph."));
        doc.FirstSection?.Body?.InsertBefore(newParagraph, originalParagraph);

        // After insertion, locate the comment again and verify its parent paragraph.
        Comment? updatedComment = doc.GetChildNodes(NodeType.Comment, true)
                                    .OfType<Comment>()
                                    .FirstOrDefault(c => c.Id == originalCommentId);

        if (updatedComment == null)
        {
            Console.WriteLine("Comment disappeared after insertion.");
            return;
        }

        // The comment's ID should remain unchanged.
        bool idUnchanged = updatedComment.Id == originalCommentId;

        // The comment should still be attached to the original paragraph.
        bool parentCorrect = updatedComment.ParentNode is Paragraph parentPara &&
                             parentPara.GetText().Contains("Original paragraph with comment.");

        // Output validation results.
        Console.WriteLine($"Comment ID unchanged: {idUnchanged}");
        Console.WriteLine($"Comment attached to correct paragraph after insertion: {parentCorrect}");

        // Save the modified document.
        const string modifiedFile = "modified.docx";
        doc.Save(modifiedFile);
    }
}
