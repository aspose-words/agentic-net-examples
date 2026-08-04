using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCommentExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add a paragraph.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("This is a paragraph that will have a comment attached.");

            // Create a comment, set its author and text.
            Comment comment = new Comment(doc, "John Doe", "JD", DateTime.Now);
            comment.SetText("This is the comment text.");

            // Attach the comment to the current paragraph.
            builder.CurrentParagraph.AppendChild(comment);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string docPath = Path.Combine(outputDir, "CommentExample.docx");
            doc.Save(docPath);

            // Retrieve the first (and only) comment from the document.
            Comment retrievedComment = (Comment)doc.GetChildNodes(NodeType.Comment, true)[0];

            // Get the comment's text.
            string commentText = retrievedComment.GetText().Trim();

            // Output the retrieved comment text.
            Console.WriteLine($"Retrieved comment text: \"{commentText}\"");
        }
    }
}
