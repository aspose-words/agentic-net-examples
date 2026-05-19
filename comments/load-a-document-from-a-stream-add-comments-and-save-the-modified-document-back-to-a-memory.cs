using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph.");
        builder.Writeln("Second paragraph.");

        // Save the document to a memory stream.
        using (MemoryStream inputStream = new MemoryStream())
        {
            doc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0; // Reset for reading.

            // Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Add a comment to the first paragraph.
            Paragraph? targetParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
            if (targetParagraph != null)
            {
                // Create a new comment with author metadata.
                Comment comment = new Comment(loadedDoc, "Alice", "A", DateTime.Now);

                // Ensure the comment contains visible text.
                comment.AppendChild(new Paragraph(loadedDoc));
                comment.FirstParagraph?.AppendChild(new Run(loadedDoc, "Please review this paragraph."));

                // Attach the comment to the paragraph.
                targetParagraph.AppendChild(comment);
            }

            // Save the modified document to another memory stream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0; // Reset for any further processing.

                // Verify that the comment was added.
                var comments = loadedDoc.GetChildNodes(NodeType.Comment, true)
                                        .OfType<Comment>()
                                        .ToList();

                Console.WriteLine($"Number of comments in the document: {comments.Count}");
                foreach (Comment c in comments)
                {
                    Console.WriteLine($"Author: {c.Author}, Text: {c.GetText().Trim()}");
                }
            }
        }
    }
}
