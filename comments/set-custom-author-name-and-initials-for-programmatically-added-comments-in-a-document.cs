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
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add some content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a paragraph that will have a comment attached.");

            // Create a comment with custom author name and initials.
            Comment comment = new Comment(doc, "Alice Johnson", "AJ", DateTime.Now);
            comment.SetText("Please review this paragraph for accuracy.");

            // Attach the comment to the current paragraph.
            Paragraph? currentParagraph = builder.CurrentParagraph;
            if (currentParagraph != null)
            {
                currentParagraph.AppendChild(comment);
            }

            // Save the document to the working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CommentsSample.docx");
            doc.Save(outputPath);

            // Enumerate all comments in the document and display their metadata.
            var comments = doc.GetChildNodes(NodeType.Comment, true)
                              .OfType<Comment>()
                              .ToList();

            Console.WriteLine($"Total comments: {comments.Count}");
            foreach (Comment c in comments)
            {
                Console.WriteLine($"Author: {c.Author}, Initials: {c.Initial}, Date: {c.DateTime}");
                Console.WriteLine($"Comment text: {c.GetText().Trim()}");
                Console.WriteLine();
            }
        }
    }
}
