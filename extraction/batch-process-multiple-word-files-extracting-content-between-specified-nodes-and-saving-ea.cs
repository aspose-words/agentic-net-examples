using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class BatchExtractToPdf
{
    public static void Main()
    {
        // Prepare input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample documents.
        CreateSampleDocument(Path.Combine(inputDir, "Doc1.docx"), "Content of document 1 between bookmarks.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc2.docx"), "Another piece of text for document 2.");
        CreateSampleDocument(Path.Combine(inputDir, "Doc3.docx"), "Third document's bookmarked content.");

        // Process each document in the input folder.
        foreach (string filePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document sourceDoc = new Document(filePath);

            // Locate the bookmark that defines the extraction range.
            Bookmark rangeBookmark = sourceDoc.Range.Bookmarks["Range"];
            if (rangeBookmark == null)
                throw new InvalidOperationException($"Bookmark 'Range' not found in {Path.GetFileName(filePath)}.");

            // Determine the start and end paragraphs of the bookmarked range.
            Paragraph startParagraph = rangeBookmark.BookmarkStart.ParentNode as Paragraph;
            Paragraph endParagraph = rangeBookmark.BookmarkEnd.ParentNode as Paragraph;
            if (startParagraph == null || endParagraph == null)
                throw new InvalidOperationException("Bookmark boundaries are not valid paragraphs.");

            // Create a new document that will hold the extracted content.
            Document extractedDoc = new Document();
            extractedDoc.RemoveAllChildren();

            // Build a minimal document structure: Section -> Body.
            Section section = new Section(extractedDoc);
            extractedDoc.AppendChild(section);
            Body body = new Body(extractedDoc);
            section.AppendChild(body);

            // Import the paragraphs from the source document into the new document.
            NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
            Paragraph current = startParagraph;
            while (true)
            {
                Node importedNode = importer.ImportNode(current, true);
                body.AppendChild(importedNode);

                if (current == endParagraph)
                    break;

                // Move to the next sibling paragraph.
                current = current.NextSibling as Paragraph;
                if (current == null)
                    break; // Safety check.
            }

            // Save the extracted content as a PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(filePath) + "_Extracted.pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            extractedDoc.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF for {Path.GetFileName(filePath)}.");
        }
    }

    // Helper method to create a sample document with a bookmark named "Range".
    private static void CreateSampleDocument(string filePath, string bookmarkedContent)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document header text.");

        // Define the range to be extracted using a bookmark.
        builder.StartBookmark("Range");
        builder.Writeln(bookmarkedContent);
        builder.EndBookmark("Range");

        builder.Writeln("Document footer text.");

        doc.Save(filePath);
    }
}
