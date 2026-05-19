using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputFolder);

        // Paths for the original and revised documents.
        string originalPath = Path.Combine(outputFolder, "original.docx");
        string revisedPath = Path.Combine(outputFolder, "revised.docx");

        // -----------------------------------------------------------------
        // 1. Create a sample document with comments from different authors.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice.
        builder.Writeln("This is the first paragraph.");
        Comment aliceComment = new Comment(doc, "Alice", "A", DateTime.Now);
        // Append the comment to the current paragraph.
        builder.CurrentParagraph?.AppendChild(aliceComment);
        // Move the builder inside the comment story and add comment text.
        builder.MoveTo(aliceComment.AppendChild(new Paragraph(doc)));
        builder.Write("Alice's comment.");

        // Second paragraph with a comment from Bob.
        builder.Writeln("This is the second paragraph.");
        Comment bobComment = new Comment(doc, "Bob", "B", DateTime.Now);
        builder.CurrentParagraph?.AppendChild(bobComment);
        builder.MoveTo(bobComment.AppendChild(new Paragraph(doc)));
        builder.Write("Bob's comment.");

        // Save the original document.
        doc.Save(originalPath);

        // ---------------------------------------------------------------
        // 2. Load the document and filter comments based on author name.
        // ---------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Enumerate all comment nodes safely.
        var allComments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                   .OfType<Comment>()
                                   .ToList();

        const string acceptedAuthor = "Alice";

        // Identify comments that should be removed (author does not match).
        var commentsToRemove = allComments
            .Where(c => !string.Equals(c.Author, acceptedAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove the unwanted comments.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the revised document containing only accepted comments.
        loadedDoc.Save(revisedPath);

        // ---------------------------------------------------------------
        // 3. Optional: Write a simple console report of remaining comments.
        // ---------------------------------------------------------------
        var remainingComments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .ToList();

        Console.WriteLine("Comments kept in the revised document:");
        foreach (Comment comment in remainingComments)
        {
            string text = comment.GetText().Trim();
            Console.WriteLine($"- Author: {comment.Author}, Text: \"{text}\"");
        }
    }
}
