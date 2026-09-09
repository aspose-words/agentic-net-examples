using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the sample input and output documents.
        const string inputPath = "input.docx";
        const string outputPath = "output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with a comment and save it.
        // -----------------------------------------------------------------
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);

        // Add a paragraph of text.
        builder.Writeln("This is a sample paragraph with a comment.");

        // Create a comment with a mixed‑case author name.
        Comment comment = new Comment(sampleDoc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Initial comment text.");
        // Append the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the document that will be loaded later.
        sampleDoc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, convert comment authors to uppercase.
        // -----------------------------------------------------------------
        Document doc = new Document(inputPath);

        // Enumerate all comment nodes safely.
        var comments = doc.GetChildNodes(NodeType.Comment, true)
                          .OfType<Comment>()
                          .ToList();

        foreach (Comment c in comments)
        {
            // Transform the author name to uppercase while preserving other metadata.
            if (!string.IsNullOrEmpty(c.Author))
            {
                c.Author = c.Author.ToUpperInvariant();
            }
        }

        // -----------------------------------------------------------------
        // Step 3: Save the modified document.
        // -----------------------------------------------------------------
        doc.Save(outputPath);
    }
}
