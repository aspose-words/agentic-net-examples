using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace CommentExtractionExample
{
    class Program
    {
        static void Main()
        {
            // Load the DOCX file using the Document(string) constructor.
            Document doc = new Document("InputDocument.docx");

            // Extract all comments.
            List<string> allComments = ExtractComments(doc);

            // Output all comments to the console.
            Console.WriteLine("All comments in the document:");
            foreach (string comment in allComments)
                Console.WriteLine("- " + comment);

            // Extract comments authored by a specific person (e.g., "John Doe").
            string targetAuthor = "John Doe";
            List<string> authorComments = ExtractCommentsByAuthor(doc, targetAuthor);

            // Output filtered comments.
            Console.WriteLine($"\nComments authored by \"{targetAuthor}\":");
            foreach (string comment in authorComments)
                Console.WriteLine("- " + comment);
        }

        // Returns the text of every Comment node in the document.
        private static List<string> ExtractComments(Document doc)
        {
            List<string> comments = new List<string>();

            // Get all comment nodes in the document (including those in headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in commentNodes)
            {
                // The comment's visible text is obtained via its GetText method.
                comments.Add(comment.GetText().Trim());
            }

            return comments;
        }

        // Returns the text of Comment nodes whose Author property matches the supplied name.
        private static List<string> ExtractCommentsByAuthor(Document doc, string author)
        {
            List<string> comments = new List<string>();

            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            foreach (Comment comment in commentNodes)
            {
                if (string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                {
                    comments.Add(comment.GetText().Trim());
                }
            }

            return comments;
        }
    }
}
