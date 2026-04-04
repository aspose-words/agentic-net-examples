using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First paragraph with a comment from Alice (to be accepted).
        builder.Writeln("First paragraph.");
        Comment commentAlice = new Comment(doc, "Alice", "A", DateTime.Now);
        commentAlice.SetText("Alice's comment.");
        builder.CurrentParagraph?.AppendChild(commentAlice);

        // Second paragraph with a comment from Bob (to be rejected).
        builder.Writeln("Second paragraph.");
        Comment commentBob = new Comment(doc, "Bob", "B", DateTime.Now);
        commentBob.SetText("Bob's comment.");
        builder.CurrentParagraph?.AppendChild(commentBob);

        // Third paragraph with a comment from Charlie (to be rejected).
        builder.Writeln("Third paragraph.");
        Comment commentCharlie = new Comment(doc, "Charlie", "C", DateTime.Now);
        commentCharlie.SetText("Charlie's comment.");
        builder.CurrentParagraph?.AppendChild(commentCharlie);

        // Save the original document.
        const string originalPath = "original.docx";
        doc.Save(originalPath);

        // Load the document for processing.
        Document processedDoc = new Document(originalPath);

        // Author whose comments we want to keep.
        const string acceptedAuthor = "Alice";

        // Get a safe copy of all comment nodes.
        var allComments = processedDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in allComments)
        {
            if (!string.Equals(c.Author, acceptedAuthor, StringComparison.OrdinalIgnoreCase))
            {
                // Remove comments from other authors.
                c.Remove();
            }
            else
            {
                // Mark accepted comments as done (optional).
                c.Done = true;
            }
        }

        // Save the revised document.
        const string revisedPath = "revised.docx";
        processedDoc.Save(revisedPath);
    }
}
