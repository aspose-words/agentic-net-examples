using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace CommentAuthorUppercase
{
    public class Program
    {
        public static void Main()
        {
            // Input and output file names.
            const string inputFile = "input.docx";
            const string outputFile = "output.docx";

            // -------------------------------------------------
            // Step 1: Create a sample document with a comment.
            // -------------------------------------------------
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);

            // Add a paragraph that will contain the comment.
            builder.Writeln("This is a sample paragraph with a comment.");

            // Create a comment with a mixed‑case author name.
            // The constructor (Document, author, initial, date) is the recommended way.
            Comment comment = new Comment(sampleDoc, "John Doe", "JD", DateTime.Now);
            // Set the comment text – this method creates a paragraph and a run inside the comment.
            comment.SetText("Initial comment text.");

            // Attach the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Save the sample document to be loaded later.
            sampleDoc.Save(inputFile);

            // -------------------------------------------------
            // Step 2: Load the document and convert comment authors to uppercase.
            // -------------------------------------------------
            Document doc = new Document(inputFile);

            // Enumerate all comment nodes safely.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            foreach (Comment c in comments)
            {
                // Guard against null (Author is never null, but we check for safety).
                if (!string.IsNullOrEmpty(c.Author))
                {
                    c.Author = c.Author.ToUpperInvariant();
                }
            }

            // -------------------------------------------------
            // Step 3: Save the modified document.
            // -------------------------------------------------
            doc.Save(outputFile);
        }
    }
}
