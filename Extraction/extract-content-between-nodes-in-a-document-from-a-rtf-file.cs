using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source RTF file.
            string rtfPath = @"C:\Docs\SourceDocument.rtf";

            // Names of the bookmarks that mark the start and end of the desired range.
            string startBookmarkName = "Start";
            string endBookmarkName = "End";

            // Extract the text between the two bookmarks.
            string extractedText = ExtractTextBetweenBookmarks(rtfPath, startBookmarkName, endBookmarkName);

            // Output the extracted text to the console.
            Console.WriteLine("Extracted Text:");
            Console.WriteLine(extractedText);

            // Optionally, save the extracted text to a plain‑text file.
            string outputPath = @"C:\Docs\ExtractedContent.txt";
            File.WriteAllText(outputPath, extractedText);
        }

        /// <summary>
        /// Loads an RTF document, locates two bookmarks, and returns the concatenated text of all nodes
        /// that lie between the start bookmark and the end bookmark (inclusive).
        /// </summary>
        /// <param name="filePath">Full path to the RTF file.</param>
        /// <param name="startBookmark">Name of the bookmark that marks the beginning of the range.</param>
        /// <param name="endBookmark">Name of the bookmark that marks the end of the range.</param>
        /// <returns>Plain text contained between the two bookmarks.</returns>
        static string ExtractTextBetweenBookmarks(string filePath, string startBookmark, string endBookmark)
        {
            // Load the RTF document using RtfLoadOptions (lifecycle rule: load).
            var loadOptions = new RtfLoadOptions();
            Document doc = new Document(filePath, loadOptions);

            // Retrieve the bookmark objects.
            Bookmark start = doc.Range.Bookmarks[startBookmark];
            Bookmark end = doc.Range.Bookmarks[endBookmark];

            // Validate that both bookmarks exist.
            if (start == null)
                throw new ArgumentException($"Bookmark '{startBookmark}' not found.");
            if (end == null)
                throw new ArgumentException($"Bookmark '{endBookmark}' not found.");

            // Use the underlying BookmarkStart/BookmarkEnd nodes.
            Node startNode = start.BookmarkStart;
            Node endNode = end.BookmarkEnd;

            // Collect all nodes in document order.
            NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);

            var sb = new StringBuilder();
            bool capture = false;

            foreach (Node node in allNodes)
            {
                // Begin capturing when we reach the start node.
                if (node == startNode)
                    capture = true;

                // If we are within the range, append the node's text.
                if (capture)
                    sb.Append(node.GetText());

                // Stop after processing the end node.
                if (node == endNode)
                    break;
            }

            return sb.ToString();
        }
    }
}
