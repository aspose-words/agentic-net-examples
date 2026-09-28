using System;
using System.IO;
using System.Linq;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a new document and add a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the first paragraph of the document.");

        // Helper to add a comment to the first paragraph.
        void AddComment(string author, string initial, string text)
        {
            Comment comment = new Comment(doc)
            {
                Author = author,
                Initial = initial,
                DateTime = DateTime.Now
            };
            // Every comment must contain at least one paragraph.
            comment.AppendChild(new Paragraph(doc));
            // Add the comment text.
            comment.FirstParagraph?.AppendChild(new Run(doc, text));

            // Attach the comment to the first paragraph if it exists.
            Paragraph? firstPara = doc.FirstSection?.Body?.FirstParagraph;
            if (firstPara != null)
            {
                firstPara.AppendChild(comment);
            }
        }

        // Add comments authored by John and Alice.
        AddComment("John", "J", "John's comment on the paragraph.");
        AddComment("Alice", "A", "Alice's comment on the paragraph.");

        // Save the original document.
        string originalPath = Path.Combine(outputDir, "original.docx");
        doc.Save(originalPath);

        // Load the document back.
        Document loadedDoc = new Document(originalPath);

        // Author whose comments should be removed.
        string targetAuthor = "John";

        // Find comments by the target author.
        var commentsToRemove = loadedDoc.GetChildNodes(NodeType.Comment, true)
            .OfType<Comment>()
            .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Remove the matching comments.
        foreach (Comment c in commentsToRemove)
        {
            c.Remove();
        }

        // Save the modified document.
        string modifiedPath = Path.Combine(outputDir, "modified.docx");
        loadedDoc.Save(modifiedPath);
    }
}
