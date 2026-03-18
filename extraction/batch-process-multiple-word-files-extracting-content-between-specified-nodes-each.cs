using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchWordExtraction
{
    class Program
    {
        static void Main()
        {
            // Folder that contains the source Word documents.
            string sourceFolder = @"C:\InputDocs";

            // Folder where the extracted PDFs will be written.
            string outputFolder = @"C:\OutputPdfs";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Names of the bookmarks that delimit the region to extract.
            const string startBookmarkName = "Start";
            const string endBookmarkName = "End";

            // Process every .docx file in the source folder.
            foreach (string filePath in Directory.GetFiles(sourceFolder, "*.docx"))
            {
                // Load the source document (lifecycle rule: load).
                Document sourceDoc = new Document(filePath);

                // Retrieve the start and end bookmarks.
                Bookmark startBookmark = sourceDoc.Range.Bookmarks[startBookmarkName];
                Bookmark endBookmark = sourceDoc.Range.Bookmarks[endBookmarkName];

                // If either bookmark is missing, skip this file.
                if (startBookmark == null || endBookmark == null)
                    continue;

                // Create a new blank document that will hold the extracted content
                // (lifecycle rule: create). The blank document already contains a
                // single section with a body and a paragraph.
                Document extractedDoc = new Document();

                // Prepare a NodeImporter to copy nodes from the source document
                // into the new document while preserving the original formatting.
                NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

                // The content between the two bookmarks is traversed node by node.
                // We start with the node immediately after the start bookmark and
                // stop when we reach the end bookmark.
                Node currentNode = startBookmark.BookmarkStart.NextSibling;

                while (currentNode != null && currentNode != endBookmark.BookmarkEnd)
                {
                    // Import the current node (including its children) into the target document.
                    Node importedNode = importer.ImportNode(currentNode, true);

                    // Append the imported node to the body of the extracted document.
                    extractedDoc.FirstSection.Body.AppendChild(importedNode);

                    // Move to the next sibling in the source document.
                    currentNode = currentNode.NextSibling;
                }

                // Build the output PDF file name.
                string outputPdfPath = Path.Combine(
                    outputFolder,
                    Path.GetFileNameWithoutExtension(filePath) + "_Extract.pdf");

                // Save the extracted document as PDF (lifecycle rule: save).
                extractedDoc.Save(outputPdfPath);
            }
        }
    }
}
