using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple document in memory.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("Hello World! This is the original document.");

        // Step 2: Save the document to a memory stream (input stream).
        using (MemoryStream inputStream = new MemoryStream())
        {
            originalDoc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Step 4: Add a comment to the first paragraph.
            // Ensure the document has at least one paragraph.
            Paragraph? targetParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph != null)
            {
                // Create a new comment with author metadata.
                Comment comment = new Comment(loadedDoc, "Alice Example", "AE", DateTime.Now);
                // Set the visible text of the comment.
                comment.SetText("Review this paragraph for clarity.");

                // Append the comment to the paragraph.
                targetParagraph.AppendChild(comment);
            }

            // Step 5: Save the modified document to another memory stream (output stream).
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset if further processing is needed.

                // For demonstration purposes, write the size of the resulting document.
                Console.WriteLine($"Modified document size: {outputStream.Length} bytes");
            }
        }
    }
}
