using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a simple document in memory.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample paragraph for comment demonstration.");

        // -----------------------------------------------------------------
        // 2. Save the document to a memory stream (acts as the input source).
        // -----------------------------------------------------------------
        using MemoryStream inputStream = new MemoryStream();
        doc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset for reading.

        // -----------------------------------------------------------------
        // 3. Load the document from the memory stream.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputStream);

        // -----------------------------------------------------------------
        // 4. Add a comment to the first paragraph, if it exists.
        // -----------------------------------------------------------------
        Paragraph? firstParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
        if (firstParagraph != null)
        {
            // Create the comment node and set its metadata.
            Comment comment = new Comment(loadedDoc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };

            // A comment must contain at least one paragraph and one run.
            Paragraph commentParagraph = new Paragraph(loadedDoc);
            comment.AppendChild(commentParagraph);
            commentParagraph.AppendChild(new Run(loadedDoc, "Please review this paragraph."));

            // Attach the comment to the target paragraph.
            firstParagraph.AppendChild(comment);
        }

        // -----------------------------------------------------------------
        // 5. Enumerate all comments to demonstrate retrieval.
        // -----------------------------------------------------------------
        var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                .OfType<Comment>()
                                .ToList();

        foreach (Comment c in comments)
        {
            Console.WriteLine($"{c.Author}: {c.GetText().Trim()}");
        }

        // -----------------------------------------------------------------
        // 6. Save the modified document to an output memory stream.
        // -----------------------------------------------------------------
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);
        outputStream.Position = 0; // Reset if further processing is needed.

        // -----------------------------------------------------------------
        // 7. Write the output stream to a deterministic file.
        // -----------------------------------------------------------------
        File.WriteAllBytes("ModifiedDocument.docx", outputStream.ToArray());
    }
}
