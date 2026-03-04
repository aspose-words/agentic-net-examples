using System;
using Aspose.Words;

namespace AsposeWordsCommentRemoval
{
    public class CommentRemover
    {
        /// <summary>
        /// Removes comments from a DOCX file.
        /// If <paramref name="author"/> is null, all comments are removed.
        /// Otherwise only comments authored by the specified user are removed.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOCX file.</param>
        /// <param name="outputPath">Full path where the modified DOCX will be saved.</param>
        /// <param name="author">Optional author name to filter comments. Pass null to delete all.</param>
        public static void RemoveComments(string inputPath, string outputPath, string? author = null)
        {
            // Load the document from the file system.
            Document doc = new Document(inputPath);

            // Retrieve all comment nodes in the document (including those in headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            // Iterate backwards so that removal does not affect the collection indexing.
            for (int i = commentNodes.Count - 1; i >= 0; i--)
            {
                Comment comment = (Comment)commentNodes[i];

                // If an author filter is supplied, remove only matching comments.
                // Otherwise remove every comment.
                if (author == null || string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                {
                    comment.Remove();
                }
            }

            // Save the modified document to the specified output path.
            doc.Save(outputPath);
        }

        // Example usage.
        public static void Main()
        {
            string sourceFile = @"C:\Docs\Sample.docx";
            string resultAll = @"C:\Docs\Sample_NoComments.docx";
            string resultAuthor = @"C:\Docs\Sample_NoJohnComments.docx";

            // Remove all comments.
            RemoveComments(sourceFile, resultAll);

            // Remove only comments authored by "John Doe".
            RemoveComments(sourceFile, resultAuthor, "John Doe");
        }
    }
}
