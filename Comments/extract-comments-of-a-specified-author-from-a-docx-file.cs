using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;

namespace AsposeWordsExamples
{
    public class CommentExtractor
    {
        /// <summary>
        /// Loads a DOCX file and returns the text of all comments authored by the specified user.
        /// </summary>
        /// <param name="docxPath">Full path to the DOCX file.</param>
        /// <param name="authorName">Exact author name to filter comments.</param>
        /// <returns>List of comment texts written by the given author.</returns>
        public static List<string> ExtractCommentsByAuthor(string docxPath, string authorName)
        {
            // Load the document from the file system.
            Document doc = new Document(docxPath);

            // Retrieve all comment nodes in the document (including those inside footnotes, headers, etc.).
            NodeCollection allComments = doc.GetChildNodes(NodeType.Comment, true);

            // Filter comments by the Author property and collect their plain text.
            List<string> authorComments = allComments
                .OfType<Comment>()
                .Where(c => string.Equals(c.Author, authorName, StringComparison.Ordinal))
                .Select(c => c.GetText().Trim())
                .ToList();

            return authorComments;
        }

        // Example usage.
        public static void Main()
        {
            string filePath = @"C:\Docs\SampleDocument.docx";
            string targetAuthor = "John Doe";

            List<string> comments = ExtractCommentsByAuthor(filePath, targetAuthor);

            Console.WriteLine($"Comments authored by \"{targetAuthor}\":");
            foreach (string commentText in comments)
            {
                Console.WriteLine($"- {commentText}");
            }
        }
    }
}
