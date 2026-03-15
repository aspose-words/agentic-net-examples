using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AsposeWordsExamples
{
    public static class CommentExtractor
    {
        /// <summary>
        /// Loads a DOCX file and extracts the text of all comments.
        /// If <paramref name="author"/> is provided, only comments authored by that person are returned.
        /// </summary>
        /// <param name="docxPath">Full path to the DOCX file.</param>
        /// <param name="author">Optional author name to filter comments. Pass null to retrieve all comments.</param>
        /// <returns>List of comment texts.</returns>
        public static List<string> ExtractComments(string docxPath, string? author = null)
        {
            // Load the document using the constructor that accepts a file name.
            Document doc = new Document(docxPath);

            // Retrieve all comment nodes in the document (including those inside headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            List<string> result = new List<string>();

            // Iterate through each comment node.
            foreach (Comment comment in commentNodes)
            {
                // If an author filter is supplied, skip comments that do not match.
                if (!string.IsNullOrEmpty(author) &&
                    !string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Get the plain text of the comment (including all its child paragraphs).
                string commentText = comment.GetText().Trim();

                // Add to the result list.
                result.Add(commentText);
            }

            return result;
        }

        // Example usage.
        public static void Main()
        {
            string filePath = @"C:\Docs\SampleDocument.docx";

            // Extract all comments.
            List<string> allComments = ExtractComments(filePath);
            Console.WriteLine("All comments:");
            foreach (string txt in allComments)
                Console.WriteLine("- " + txt);

            // Extract only comments authored by "John Doe".
            List<string> johnComments = ExtractComments(filePath, "John Doe");
            Console.WriteLine("\nComments by John Doe:");
            foreach (string txt in johnComments)
                Console.WriteLine("- " + txt);
        }
    }
}
