using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class DeleteCommentsByAuthor
{
    public static void Main()
    {
        // Define output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the sample and result documents.
        string samplePath = Path.Combine(outputDir, "sample.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // Create a sample document with two comments from different authors.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice.
        builder.Writeln("First paragraph.");
        Comment aliceComment = new Comment(doc, "Alice", "A", DateTime.Now);
        aliceComment.SetText("Comment authored by Alice.");
        builder.CurrentParagraph.AppendChild(aliceComment);

        // Second paragraph with a comment from Bob.
        builder.Writeln("Second paragraph.");
        Comment bobComment = new Comment(doc, "Bob", "B", DateTime.Now);
        bobComment.SetText("Comment authored by Bob.");
        builder.CurrentParagraph.AppendChild(bobComment);

        // Save the sample document.
        doc.Save(samplePath);

        // Load the document we just created.
        Document loadedDoc = new Document(samplePath);

        // Author whose comments should be removed.
        const string targetAuthor = "Alice";

        // Find all comments authored by the target author (case‑insensitive).
        var commentsToRemove = loadedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove each matching comment safely.
        foreach (Comment comment in commentsToRemove)
        {
            comment.Remove();
        }

        // Save the document after removal.
        loadedDoc.Save(resultPath);
    }
}
