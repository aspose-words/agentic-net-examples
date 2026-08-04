using System;
using System.IO;
using System.Linq;
using Aspose.Words;

namespace CommentProcessingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with comments from different authors.
            Document originalDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(originalDoc);

            // First paragraph with a comment from Alice.
            builder.Writeln("This is the first paragraph.");
            Comment aliceComment = new Comment(originalDoc, "Alice", "AL", DateTime.Now);
            aliceComment.SetText("Alice's review note.");
            builder.CurrentParagraph.AppendChild(aliceComment);

            // Second paragraph with a comment from Bob.
            builder.Writeln("This is the second paragraph.");
            Comment bobComment = new Comment(originalDoc, "Bob", "BO", DateTime.Now);
            bobComment.SetText("Bob's suggestion.");
            builder.CurrentParagraph.AppendChild(bobComment);

            // Save the original document.
            const string originalPath = "original.docx";
            originalDoc.Save(originalPath);

            // Load the document for processing.
            Document processedDoc = new Document(originalPath);

            // Define the author whose comments we want to keep.
            const string allowedAuthor = "Alice";

            // Find comments that are NOT from the allowed author.
            var commentsToRemove = processedDoc
                .GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .Where(c => !string.Equals(c.Author, allowedAuthor, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Remove the unwanted comments safely.
            foreach (Comment comment in commentsToRemove)
            {
                comment.Remove();
            }

            // Save the revised document.
            const string revisedPath = "revised.docx";
            processedDoc.Save(revisedPath);

            // Report the results.
            Console.WriteLine($"Original document saved as: {Path.GetFullPath(originalPath)}");
            Console.WriteLine($"Revised document saved as: {Path.GetFullPath(revisedPath)}");
            Console.WriteLine($"Comments retained (author = {allowedAuthor}):");

            var retainedComments = processedDoc
                .GetChildNodes(NodeType.Comment, true)
                .OfType<Comment>()
                .ToList();

            foreach (Comment comment in retainedComments)
            {
                Console.WriteLine($"- Author: {comment.Author}, Text: {comment.GetText().Trim()}");
            }
        }
    }
}
