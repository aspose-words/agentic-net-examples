using System;
using System.Collections.Generic;
using Aspose.Words;

namespace AsposeWordsCommentExtraction
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            const string sourceFile = @"C:\Docs\SourceDocument.docx";

            // Path to the output DOCX file that will contain the extracted comments.
            const string outputFileAll = @"C:\Docs\ExtractedComments_All.docx";
            const string outputFileByAuthor = @"C:\Docs\ExtractedComments_JohnDoe.docx";

            // Extract all comments.
            ExtractComments(sourceFile, outputFileAll);

            // Extract only comments authored by "John Doe".
            ExtractComments(sourceFile, outputFileByAuthor, "John Doe");
        }

        /// <summary>
        /// Extracts comments from a Word document and saves them into a new document.
        /// If <paramref name="author"/> is null, all comments are extracted; otherwise only comments
        /// whose <see cref="Comment.Author"/> matches the supplied author are extracted.
        /// </summary>
        /// <param name="inputPath">Path to the source DOCX file.</param>
        /// <param name="outputPath">Path where the new document with extracted comments will be saved.</param>
        /// <param name="author">Optional author name to filter comments. Pass null to include all.</param>
        static void ExtractComments(string inputPath, string outputPath, string? author = null)
        {
            // Load the source document.
            Document sourceDoc = new Document(inputPath);

            // Retrieve all comment nodes (deep search).
            NodeCollection commentNodes = sourceDoc.GetChildNodes(NodeType.Comment, true);

            // Collect comments that satisfy the optional author filter.
            List<Comment> selectedComments = new List<Comment>();
            foreach (Comment comment in commentNodes)
            {
                if (author == null || string.Equals(comment.Author, author, StringComparison.OrdinalIgnoreCase))
                {
                    selectedComments.Add(comment);
                }
            }

            // Create a new blank document to hold the extracted comments.
            Document resultDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(resultDoc);

            if (selectedComments.Count == 0)
            {
                builder.Writeln("No comments found matching the specified criteria.");
            }
            else
            {
                foreach (Comment comment in selectedComments)
                {
                    builder.Writeln($"Author: {comment.Author}");
                    builder.Writeln($"Date: {comment.DateTime}");
                    builder.Writeln($"Done: {comment.Done}");
                    builder.Writeln("Comment Text:");

                    foreach (Paragraph para in comment.Paragraphs)
                    {
                        builder.Writeln(para.GetText().TrimEnd('\r', '\n'));
                    }

                    builder.Writeln(new string('-', 40));
                }
            }

            // Save the result document.
            resultDoc.Save(outputPath);
        }
    }
}
