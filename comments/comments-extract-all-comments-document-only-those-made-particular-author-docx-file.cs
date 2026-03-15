using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

namespace CommentExtractionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string sourceFile = @"C:\Docs\SourceDocument.docx";

            // Path to the destination DOCX file that will contain the extracted comments.
            const string destinationFileAll = @"C:\Docs\ExtractedComments_All.docx";

            // Path to the destination DOCX file that will contain comments only from a specific author.
            const string destinationFileAuthor = @"C:\Docs\ExtractedComments_JohnDoe.docx";

            // Extract all comments.
            ExtractComments(sourceFile, destinationFileAll);

            // Extract only comments authored by "John Doe".
            ExtractComments(sourceFile, destinationFileAuthor, "John Doe");
        }

        /// <summary>
        /// Loads a Word document, extracts its comments (optionally filtered by author),
        /// and saves the extracted comments into a new document.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOCX file.</param>
        /// <param name="outputPath">Full path where the new document with extracted comments will be saved.</param>
        /// <param name="authorFilter">If not null, only comments whose Author matches this value are extracted.</param>
        static void ExtractComments(string inputPath, string outputPath, string? authorFilter = null)
        {
            // Load the source document.
            Document sourceDoc = new Document(inputPath);

            // Retrieve all comment nodes (including those inside nested structures).
            NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

            // Collect comments that satisfy the optional author filter.
            List<Comment> selectedComments = new List<Comment>();
            foreach (Comment comment in commentNodes)
            {
                if (authorFilter != null &&
                    !string.Equals(comment.Author, authorFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Skip non‑matching comments.
                }
                selectedComments.Add(comment);
            }

            // Create a new blank document to hold the extracted comments.
            Document resultDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(resultDoc);

            if (selectedComments.Count == 0)
            {
                builder.Writeln("No comments were found matching the specified criteria.");
            }
            else
            {
                foreach (Comment comment in selectedComments)
                {
                    string header = $"Author: {comment.Author ?? "Unknown"} | Date: {comment.DateTime:g}";
                    builder.Writeln(header);
                    builder.Writeln(comment.GetText().Trim());
                    builder.Writeln(); // Blank line between comments.
                }
            }

            // Save the result document.
            resultDoc.Save(outputPath);
        }
    }
}
