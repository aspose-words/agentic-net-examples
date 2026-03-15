using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace CommentExtractionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string sourceFile = @"C:\Docs\SourceDocument.docx";

            // Path to the output DOCX file that will contain the extracted comments.
            const string outputFileAll = @"C:\Docs\ExtractedComments_All.docx";

            // Path to the output DOCX file that will contain only comments authored by "John Doe".
            const string outputFileFiltered = @"C:\Docs\ExtractedComments_JohnDoe.docx";

            // Extract all comments.
            ExtractComments(sourceFile, outputFileAll);

            // Extract only comments authored by "John Doe".
            ExtractComments(sourceFile, outputFileFiltered, "John Doe");
        }

        /// <summary>
        /// Extracts comments from a DOCX file and writes them to a new document.
        /// If <paramref name="authorFilter"/> is null, all comments are extracted;
        /// otherwise only comments whose <see cref="Comment.Author"/> matches the filter are extracted.
        /// </summary>
        /// <param name="inputPath">Path to the source DOCX file.</param>
        /// <param name="outputPath">Path where the new document with extracted comments will be saved.</param>
        /// <param name="authorFilter">Optional author name to filter comments. Pass null to include all comments.</param>
        static void ExtractComments(string inputPath, string outputPath, string authorFilter = null)
        {
            // Load the source document using the Document(string) constructor.
            Document sourceDoc = new Document(inputPath);

            // Create a new blank document that will hold the extracted comments.
            Document resultDoc = new Document();

            // Use DocumentBuilder to write comment information into the result document.
            DocumentBuilder builder = new DocumentBuilder(resultDoc);

            // Retrieve all Comment nodes from the source document.
            NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

            // Prepare a list to hold the comments that satisfy the filter criteria.
            List<Comment> matchingComments = new List<Comment>();

            foreach (Comment comment in commentNodes)
            {
                // If an author filter is supplied, compare it (case‑insensitive) with the comment's Author.
                if (authorFilter == null ||
                    string.Equals(comment.Author, authorFilter, StringComparison.OrdinalIgnoreCase))
                {
                    matchingComments.Add(comment);
                }
            }

            // If no comments match, write a placeholder message.
            if (matchingComments.Count == 0)
            {
                builder.Writeln(authorFilter == null
                    ? "No comments were found in the document."
                    : $"No comments authored by \"{authorFilter}\" were found.");
            }
            else
            {
                // Write each matching comment to the result document.
                foreach (Comment comment in matchingComments)
                {
                    // Write author name.
                    builder.Font.Bold = true;
                    builder.Writeln($"Author: {comment.Author}");

                    // Write comment date (if needed).
                    builder.Font.Bold = false;
                    builder.Writeln($"Date: {comment.DateTime}");

                    // Write the comment's text. Comment.GetText() returns the full text of the comment.
                    builder.Writeln("Comment:");
                    builder.Writeln(comment.GetText().Trim());

                    // Insert a horizontal line separator between comments for readability.
                    builder.InsertHorizontalRule();
                }
            }

            // Save the result document using the Document.Save(string) method.
            resultDoc.Save(outputPath);
        }
    }
}
