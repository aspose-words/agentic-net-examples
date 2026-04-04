using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a simple document in memory.
        Document originalDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(originalDoc);
        builder.Writeln("This is a sample paragraph that will receive a comment.");

        // Step 2: Save the document to a memory stream (input stream).
        using (MemoryStream inputStream = new MemoryStream())
        {
            originalDoc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset for reading.

            // Step 3: Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Step 4: Add a comment to the first paragraph.
            // Create the comment node with author metadata.
            Comment comment = new Comment(loadedDoc, "John Doe", "JD", DateTime.Now);
            comment.SetText("This is a comment added programmatically.");

            // Anchor the comment to a region of text using CommentRangeStart/End.
            Paragraph paragraph = loadedDoc.FirstSection.Body.FirstParagraph;
            // Insert the start of the comment range.
            paragraph.AppendChild(new CommentRangeStart(loadedDoc, comment.Id));
            // Insert some text that will be highlighted as commented.
            paragraph.AppendChild(new Run(loadedDoc, "Commented text."));
            // Insert the end of the comment range.
            paragraph.AppendChild(new CommentRangeEnd(loadedDoc, comment.Id));
            // Finally, attach the comment itself.
            paragraph.AppendChild(comment);

            // Step 5: Save the modified document to a new memory stream (output stream).
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset for any further processing.

                // Optional: Verify that the comment exists by enumerating comments.
                var comments = loadedDoc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
                foreach (Comment c in comments)
                {
                    Console.WriteLine($"Author: {c.Author}");
                    Console.WriteLine($"Date: {c.DateTime}");
                    Console.WriteLine($"Text: {c.GetText().Trim()}");
                    Console.WriteLine();
                }

                // The outputStream now contains the DOCX with the added comment.
                // It can be written to a file if needed (not required by the task).
                // Example (commented out):
                // File.WriteAllBytes("ModifiedDocument.docx", outputStream.ToArray());
            }
        }
    }
}
