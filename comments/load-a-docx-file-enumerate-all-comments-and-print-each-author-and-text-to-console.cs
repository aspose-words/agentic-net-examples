using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph.");
        Comment comment1 = new Comment(sampleDoc, "Alice", "A", DateTime.Now);
        comment1.SetText("This is Alice's comment.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph.");
        Comment comment2 = new Comment(sampleDoc, "Bob", "B", DateTime.Now);
        comment2.SetText("Bob added this comment.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the sample document to a temporary file.
        string tempFile = Path.Combine(Path.GetTempPath(), "sample.docx");
        sampleDoc.Save(tempFile);

        // Load the document from the file.
        Document loadedDoc = new Document(tempFile);

        // Enumerate all comments and print author and text.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            // Trim the comment text to remove trailing whitespace.
            string text = c.GetText()?.Trim() ?? string.Empty;
            Console.WriteLine($"{c.Author}: {text}");
        }

        // Clean up the temporary file.
        if (File.Exists(tempFile))
        {
            File.Delete(tempFile);
        }
    }
}
