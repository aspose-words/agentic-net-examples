using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Notes;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is the first paragraph of the document.");
        builder.Writeln("This is the second paragraph of the document.");

        // Save the original document to a memory stream.
        using MemoryStream inputStream = new MemoryStream();
        originalDoc.Save(inputStream, SaveFormat.Docx);
        inputStream.Position = 0; // Reset stream position for reading.

        // Load the document from the memory stream.
        Document loadedDoc = new Document(inputStream);

        // Add a comment to the first paragraph.
        Paragraph? targetParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
        if (targetParagraph != null)
        {
            // Create a new comment node.
            Comment comment = new Comment(loadedDoc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };

            // The comment must contain at least one paragraph and run.
            Paragraph commentParagraph = new Paragraph(loadedDoc);
            commentParagraph.AppendChild(new Run(loadedDoc, "Please review this paragraph."));
            comment.AppendChild(commentParagraph);

            // Append the comment to the target paragraph.
            targetParagraph.AppendChild(comment);
        }

        // Save the modified document to another memory stream.
        using MemoryStream outputStream = new MemoryStream();
        loadedDoc.Save(outputStream, SaveFormat.Docx);
        outputStream.Position = 0; // Reset for any further processing.

        // (Optional) Write the output to a file for verification.
        File.WriteAllBytes("ModifiedDocument.docx", outputStream.ToArray());
    }
}
