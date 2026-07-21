using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare file names.
        string inputPath = "input.docx";
        string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with comments and save it.
        // -----------------------------------------------------------------
        Document createDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(createDoc);

        // Add a paragraph that will contain comments.
        builder.Writeln("This is a sample paragraph with comments.");

        // First comment.
        Comment comment1 = new Comment(createDoc, "John Doe", "JD", DateTime.Now);
        comment1.SetText("First comment.");
        builder.CurrentParagraph.AppendChild(comment1);

        // Second comment.
        Comment comment2 = new Comment(createDoc, "Jane Smith", "JS", DateTime.Now);
        comment2.SetText("Second comment.");
        builder.CurrentParagraph.AppendChild(comment2);

        // Save the document that will be loaded later.
        createDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, convert comment authors to uppercase.
        // -----------------------------------------------------------------
        Document loadDoc = new Document(inputPath);

        // Enumerate all comment nodes safely.
        var comments = loadDoc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

        foreach (Comment c in comments)
        {
            // Transform the author name to uppercase.
            if (!string.IsNullOrEmpty(c.Author))
                c.Author = c.Author.ToUpperInvariant();
        }

        // -----------------------------------------------------------------
        // Step 3: Save the updated document.
        // -----------------------------------------------------------------
        loadDoc.Save(outputPath);
    }
}
