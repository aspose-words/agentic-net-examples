using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document initialDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(initialDoc);
        builder.Writeln("This is the first paragraph of the document.");
        builder.Writeln("This is the second paragraph of the document.");

        // Save the document to a memory stream (input stream).
        using (MemoryStream inputStream = new MemoryStream())
        {
            initialDoc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset for reading.

            // Load the document from the memory stream.
            Document doc = new Document(inputStream);

            // Create a comment.
            Comment comment = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };
            // Add visible text to the comment.
            comment.AppendChild(new Paragraph(doc));
            comment.FirstParagraph?.AppendChild(new Run(doc, "Please review this paragraph."));

            // Attach the comment to the first paragraph of the document.
            Paragraph? firstParagraph = doc.FirstSection?.Body?.FirstParagraph;
            if (firstParagraph != null)
            {
                firstParagraph.AppendChild(comment);
            }

            // Save the modified document to an output memory stream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                doc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset for further use.

                // Write the output to a file for verification.
                File.WriteAllBytes("ModifiedDocument.docx", outputStream.ToArray());
            }
        }
    }
}
