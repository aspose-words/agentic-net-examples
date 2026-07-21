using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello world! This paragraph will have a comment.");

        // Save the document to a memory stream.
        using (MemoryStream inputStream = new MemoryStream())
        {
            doc.Save(inputStream, SaveFormat.Docx);
            inputStream.Position = 0;

            // Load the document from the memory stream.
            Document loadedDoc = new Document(inputStream);

            // Add a comment to the first paragraph.
            Paragraph? firstParagraph = loadedDoc.FirstSection?.Body?.FirstParagraph;
            if (firstParagraph != null)
            {
                Comment comment = new Comment(loadedDoc)
                {
                    Author = "Alice",
                    Initial = "A",
                    DateTime = DateTime.Now
                };

                // Add visible text to the comment.
                comment.AppendChild(new Paragraph(loadedDoc));
                comment.FirstParagraph?.AppendChild(new Run(loadedDoc, "Review this paragraph."));

                // Append the comment to the paragraph.
                firstParagraph.AppendChild(comment);
            }

            // Save the modified document to another memory stream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                loadedDoc.Save(outputStream, SaveFormat.Docx);
                outputStream.Position = 0;

                // Enumerate comments and write their details to the console.
                var comments = loadedDoc.GetChildNodes(NodeType.Comment, true).OfType<Comment>();
                foreach (Comment c in comments)
                {
                    Console.WriteLine($"Author: {c.Author}, Text: {c.GetText().Trim()}");
                }
            }
        }
    }
}
