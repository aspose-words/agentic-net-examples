using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document initialDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(initialDoc);
        builder.Writeln("This is a sample paragraph for comment demonstration.");

        // Save the document to a memory stream.
        using (MemoryStream inputStream = new MemoryStream())
        {
            initialDoc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset stream position for reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Add a comment to the first paragraph.
            Paragraph? firstParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
            if (firstParagraph != null)
            {
                // Create a new comment with author metadata.
                Comment comment = new Comment(loadedDoc, "Alice", "A", DateTime.Now);
                // Set the comment text; this creates the required paragraph inside the comment.
                comment.SetText("Please review this paragraph.");

                // Append the comment to the paragraph.
                firstParagraph.AppendChild(comment);
            }

            // Save the modified document to another memory stream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset for any further processing.

                // Enumerate and display comments to verify.
                var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                         .OfType<Comment>()
                                         .ToList();

                foreach (Comment c in comments)
                {
                    Console.WriteLine($"Author: {c.Author}, Text: {c.GetText().Trim()}");
                }

                // The outputStream now contains the DOCX with the added comment.
                // It can be written to a file if needed (omitted as per requirements).
            }
        }
    }
}
