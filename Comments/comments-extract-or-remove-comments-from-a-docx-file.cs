using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsCommentDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths.
            string inputPath = @"C:\Docs\SampleWithComments.docx";
            string outputPath = @"C:\Docs\SampleWithoutComments.docx";

            // Load the document (lifecycle rule: load).
            Document doc = new Document(inputPath);

            // -----------------------------------------------------------------
            // Extract all comments from the document.
            // -----------------------------------------------------------------
            // Get all comment nodes in the document (including those in headers/footers).
            NodeCollection commentNodes = doc.GetChildNodes(NodeType.Comment, true);
            List<string> extractedComments = new List<string>();

            foreach (Comment comment in commentNodes.Cast<Comment>())
            {
                // The comment text is stored in the comment's range.
                // Trim to remove trailing paragraph marks.
                string commentText = comment.GetText().TrimEnd('\r', '\a');
                extractedComments.Add(commentText);
            }

            // Output extracted comments to console (or handle as needed).
            Console.WriteLine("Extracted Comments:");
            foreach (string txt in extractedComments)
                Console.WriteLine("- " + txt);

            // -----------------------------------------------------------------
            // Remove all comments from the document.
            // -----------------------------------------------------------------
            // Iterate over a copy of the collection because removing nodes modifies the collection.
            foreach (Comment comment in commentNodes.Cast<Comment>().ToList())
            {
                comment.Remove(); // Remove the comment node from its parent.
            }

            // Save the modified document (lifecycle rule: save).
            doc.Save(outputPath);
            Console.WriteLine($"Document saved without comments to: {outputPath}");
        }
    }
}
