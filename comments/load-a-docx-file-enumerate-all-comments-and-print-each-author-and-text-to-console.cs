using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the temporary document that will contain comments.
        string docPath = "SampleWithComments.docx";

        // -----------------------------------------------------------------
        // Create a new document and add a few comments to demonstrate.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);

        // First paragraph with a comment.
        builder.Writeln("First paragraph with a comment.");
        Comment comment1 = new Comment(createDoc, "Alice", "A", DateTime.Now);
        comment1.SetText("Review the first paragraph.");
        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment1);

        // Second paragraph with a comment.
        builder.Writeln("Second paragraph with another comment.");
        Comment comment2 = new Comment(createDoc, "Bob", "B", DateTime.Now);
        comment2.SetText("Check the wording here.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the document to disk.
        createDoc.Save(docPath);

        // -----------------------------------------------------------------
        // Load the document from disk.
        // -----------------------------------------------------------------
        Document loadDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // Enumerate all comments and print author and text to the console.
        // -----------------------------------------------------------------
        var comments = loadDoc
            .GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .ToList();

        foreach (Comment c in comments)
        {
            // Trim the comment text to remove any trailing whitespace or line breaks.
            string text = c.GetText().Trim();
            Console.WriteLine($"{c.Author}: {text}");
        }

        // Clean up the temporary file (optional).
        // File.Delete(docPath);
    }
}
