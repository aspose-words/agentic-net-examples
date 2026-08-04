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
        builder.Writeln("This is the original paragraph.");

        // Step 2: Save the document to a memory stream.
        using (MemoryStream inputStream = new MemoryStream())
        {
            originalDoc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Step 4: Add a comment to the first paragraph.
            Paragraph firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;
            Comment comment = new Comment(loadedDoc, "Alice", "A", DateTime.Now);
            comment.SetText("Review this paragraph.");
            firstParagraph.AppendChild(comment);

            // Optional: Enumerate comments and write to console.
            var comments = loadedDoc.GetChildNodes(NodeType.Comment, true);
            foreach (Comment c in comments.OfType<Comment>())
            {
                Console.WriteLine($"Comment by {c.Author}: {c.GetText().Trim()}");
            }

            // Step 5: Save the modified document to another memory stream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset for further use.

                // For demonstration purposes, write the output to a file.
                File.WriteAllBytes("ModifiedDocument.docx", outputStream.ToArray());
            }
        }
    }
}
