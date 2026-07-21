using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names for the sample and the result.
        string originalPath = "original.docx";
        string resultPath = "result.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with comments from different authors.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        Comment aliceComment = new Comment(doc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("Comment authored by Alice.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(aliceComment);

        // Second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        Comment bobComment = new Comment(doc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Comment authored by Bob.");
        builder.CurrentParagraph.AppendChild(bobComment);

        // Save the sample document.
        doc.Save(originalPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document and delete all comments authored by Alice.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);
        string targetAuthor = "Alice";

        // Enumerate all comment nodes, filter by author, and collect to a list to avoid
        // modifying the collection while iterating.
        var commentsToDelete = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove each matching comment from the document.
        foreach (Comment comment in commentsToDelete)
        {
            comment.Remove();
        }

        // -----------------------------------------------------------------
        // Step 3: Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(resultPath);
    }
}
