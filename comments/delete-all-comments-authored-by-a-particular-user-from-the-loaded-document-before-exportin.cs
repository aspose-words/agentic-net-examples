using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add first paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        Comment aliceComment1 = new Comment(doc, "Alice", "A", DateTime.Now);
        aliceComment1.SetText("Comment from Alice.");
        builder.CurrentParagraph.AppendChild(aliceComment1);

        // Add second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        Comment bobComment = new Comment(doc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Comment from Bob.");
        builder.CurrentParagraph.AppendChild(bobComment);

        // Add third paragraph with another comment from Alice.
        builder.Writeln("Third paragraph.");
        Comment aliceComment2 = new Comment(doc, "Alice", "A", DateTime.Now);
        aliceComment2.SetText("Another comment from Alice.");
        builder.CurrentParagraph.AppendChild(aliceComment2);

        // Define the author whose comments should be removed.
        const string targetAuthor = "Alice";

        // Find all comments authored by the target author.
        var commentsToRemove = doc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove each matching comment safely.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
