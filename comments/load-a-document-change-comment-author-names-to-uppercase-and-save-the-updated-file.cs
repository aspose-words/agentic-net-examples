using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Prepare temporary folder and file paths.
        string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Temp");
        Directory.CreateDirectory(tempFolder);
        string inputPath = Path.Combine(tempFolder, "input.docx");
        string outputPath = Path.Combine(tempFolder, "output.docx");

        // -----------------------------------------------------------------
        // Step 1: Create a sample document with a comment and save it.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a paragraph with a comment.");

        // Create a comment with a sample author.
        Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
        comment.SetText("Sample comment.");

        // Attach the comment to the current paragraph.
        builder.CurrentParagraph.AppendChild(comment);

        // Save the sample document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // Step 2: Load the document, convert comment authors to uppercase.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        // Enumerate all comment nodes safely.
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            // Guard against null or empty author values.
            if (!string.IsNullOrEmpty(c.Author))
            {
                c.Author = c.Author.ToUpperInvariant();
            }
        }

        // Save the modified document.
        loadedDoc.Save(outputPath);
    }
}
