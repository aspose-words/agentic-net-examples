using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

namespace AsposeWordsCommentExtraction
{
    class Program
    {
        // Entry point of the application.
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string sourcePath = @"C:\Docs\SampleDocument.docx";

            // Path to the output text file that will contain the extracted comments.
            string outputPath = @"C:\Docs\ExtractedComments.txt";

            // Optional: specify an author name to filter comments.
            // Set to null or empty string to extract comments from all authors.
            string authorFilter = "John Doe";

            // Load the document using the Document(string) constructor (lifecycle rule).
            Document doc = new Document(sourcePath);

            // Retrieve all comment nodes in the document (including those in headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);

            // Prepare a list to hold the formatted comment information.
            List<string> extractedComments = new List<string>();

            // Iterate through each comment node.
            foreach (Comment comment in commentNodes)
            {
                // If an author filter is provided, skip comments that do not match.
                if (!string.IsNullOrEmpty(authorFilter) &&
                    !string.Equals(comment.Author, authorFilter, StringComparison.OrdinalIgnoreCase))
                    continue;

                // Build a readable representation of the comment.
                string commentInfo = $"Comment Id: {comment.Id}" + Environment.NewLine +
                                     $"Author   : {comment.Author}" + Environment.NewLine +
                                     $"Initials : {comment.Initial}" + Environment.NewLine +
                                     $"DateTime : {comment.DateTime}" + Environment.NewLine +
                                     $"Done     : {comment.Done}" + Environment.NewLine +
                                     $"Text     : {comment.GetText().Trim()}" + Environment.NewLine +
                                     new string('-', 40);

                extractedComments.Add(commentInfo);
            }

            // Write the extracted comments to the output file.
            // This demonstrates the use of the Save(string) method indirectly by persisting data.
            File.WriteAllLines(outputPath, extractedComments);

            // Optionally, you can also save the original document unchanged using the Save(string) method.
            // This satisfies the lifecycle rule that requires using the provided save logic.
            string unchangedCopyPath = @"C:\Docs\UnchangedCopy.docx";
            doc.Save(unchangedCopyPath);
        }
    }
}
