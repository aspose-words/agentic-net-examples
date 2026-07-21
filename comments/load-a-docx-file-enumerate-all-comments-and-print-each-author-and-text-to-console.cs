using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace CommentEnumerationExample
{
    public class Program
    {
        public static void Main()
        {
            // Prepare a temporary folder for the sample document.
            string tempFolder = Path.Combine(Directory.GetCurrentDirectory(), "Temp");
            Directory.CreateDirectory(tempFolder);

            // Path of the sample DOCX file.
            string samplePath = Path.Combine(tempFolder, "SampleWithComments.docx");

            // Create a new document and add a paragraph with a comment.
            Document docToCreate = new Document();
            DocumentBuilder builder = new DocumentBuilder(docToCreate);
            builder.Writeln("This is a paragraph that will have a comment.");

            // Create a comment, set its metadata, and add text to it.
            Comment comment = new Comment(docToCreate, "Alice", "A", DateTime.Now);
            comment.SetText("Review this paragraph for clarity.");

            // Append the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Save the document to disk.
            docToCreate.Save(samplePath);

            // Load the document from the file.
            Document loadedDoc = new Document(samplePath);

            // Enumerate all comment nodes in the document.
            var comments = loadedDoc
                .GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .ToList();

            // Print author and text of each comment.
            foreach (Comment c in comments)
            {
                // GetText() returns the comment text including any paragraph breaks.
                string text = c.GetText()?.Trim() ?? string.Empty;
                Console.WriteLine($"{c.Author}: {text}");
            }

            // Optional cleanup.
            // File.Delete(samplePath);
            // Directory.Delete(tempFolder, true);
        }
    }
}
