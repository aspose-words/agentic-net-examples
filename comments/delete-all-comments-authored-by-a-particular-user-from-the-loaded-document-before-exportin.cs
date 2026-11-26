using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace CommentDeletionExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output directory.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // Step 1: Create a sample document with two comments.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // First paragraph with a comment authored by John Doe.
            builder.Writeln("This is the first paragraph.");
            Comment commentJohn = new Comment(doc)
            {
                Author = "John Doe",
                Initial = "JD",
                DateTime = DateTime.Now
            };
            commentJohn.AppendChild(new Paragraph(doc));
            commentJohn.FirstParagraph?.AppendChild(new Run(doc, "Comment from John Doe."));
            // Attach the comment to the first paragraph.
            doc.FirstSection.Body.FirstParagraph?.AppendChild(commentJohn);

            // Second paragraph with a comment authored by Jane Smith.
            builder.Writeln("This is the second paragraph.");
            Comment commentJane = new Comment(doc)
            {
                Author = "Jane Smith",
                Initial = "JS",
                DateTime = DateTime.Now
            };
            commentJane.AppendChild(new Paragraph(doc));
            commentJane.FirstParagraph?.AppendChild(new Run(doc, "Comment from Jane Smith."));
            // Attach the comment to the second paragraph.
            Paragraph? secondParagraph = doc.FirstSection.Body?.LastParagraph;
            secondParagraph?.AppendChild(commentJane);

            // Save the original document.
            string originalPath = Path.Combine(outputDir, "original.docx");
            doc.Save(originalPath);

            // -----------------------------------------------------------------
            // Step 2: Load the document from file (simulating a real load).
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(originalPath);

            // -----------------------------------------------------------------
            // Step 3: Delete all comments authored by a specific user.
            // -----------------------------------------------------------------
            string targetAuthor = "John Doe";

            var commentsToRemove = loadedDoc.GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .Where(c => string.Equals(c.Author, targetAuthor, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (Comment c in commentsToRemove)
            {
                c.Remove();
            }

            // -----------------------------------------------------------------
            // Step 4: Save the modified document.
            // -----------------------------------------------------------------
            string modifiedPath = Path.Combine(outputDir, "modified.docx");
            loadedDoc.Save(modifiedPath);
        }
    }
}
