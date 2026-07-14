using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsCommentsExample
{
    public class Program
    {
        public static void Main()
        {
            // Ensure the output directory exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph that will contain the comment.
            builder.Writeln("This is a paragraph that will have a comment.");

            // Create a comment with custom author name and initials.
            Comment comment = new Comment(doc, "Custom Author", "CA", DateTime.Now);
            comment.SetText("This is a custom comment added programmatically.");

            // Append the comment to the current paragraph.
            builder.CurrentParagraph?.AppendChild(comment);

            // Save the document.
            string docPath = Path.Combine(outputDir, "DocumentWithCustomComment.docx");
            doc.Save(docPath);

            // Enumerate all comments in the document and output their metadata.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .ToList();

            foreach (Comment c in comments)
            {
                Console.WriteLine($"Author: {c.Author}, Initial: {c.Initial}, Text: {c.GetText().Trim()}");
            }
        }
    }
}
