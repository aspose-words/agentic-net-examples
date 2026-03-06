using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace CommentExtractor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the DOCX file.
            string docPath = @"C:\Docs\Sample.docx";

            // Author whose comments we want to extract.
            string targetAuthor = "John Doe";

            // Extract comments.
            List<string> comments = ExtractCommentsByAuthor(docPath, targetAuthor);

            // Output the extracted comments.
            Console.WriteLine($"Comments by \"{targetAuthor}\":");
            foreach (string text in comments)
            {
                Console.WriteLine($"- {text}");
            }
        }

        /// <summary>
        /// Loads a document and returns the text of all comments authored by the specified person.
        /// </summary>
        /// <param name="filePath">Full path to the DOCX file.</param>
        /// <param name="author">Author name to filter comments.</param>
        /// <returns>List of comment texts.</returns>
        static List<string> ExtractCommentsByAuthor(string filePath, string author)
        {
            // Load the document using Aspose.Words Document constructor (lifecycle rule: load).
            Document doc = new Document(filePath);

            // Retrieve all comment nodes in the document (including those inside comment ranges).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            // Filter comments by the specified author and collect their trimmed text.
            List<string> result = commentNodes
                .OfType<Comment>()
                .Where(c => string.Equals(c.Author, author, StringComparison.OrdinalIgnoreCase))
                .Select(c => c.GetText().Trim())
                .ToList();

            return result;
        }
    }
}
