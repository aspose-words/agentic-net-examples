using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Newtonsoft.Json;

public class Program
{
    public static void Main(string[] args)
    {
        // Determine bookmark names from command‑line arguments or use defaults.
        string startBookmarkName = args.Length > 0 ? args[0] : "Start";
        string endBookmarkName = args.Length > 1 ? args[1] : "End";

        // -----------------------------------------------------------------
        // 1. Create a sample source document containing the two bookmarks.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph before start bookmark.");

        builder.StartBookmark(startBookmarkName);
        builder.Writeln("This is the first paragraph inside the start bookmark.");
        builder.Writeln("This is the second paragraph inside the start bookmark.");
        builder.EndBookmark(startBookmarkName);

        builder.Writeln("Paragraph between bookmarks.");

        builder.StartBookmark(endBookmarkName);
        builder.Writeln("This is the first paragraph inside the end bookmark.");
        builder.Writeln("This is the second paragraph inside the end bookmark.");
        builder.EndBookmark(endBookmarkName);

        builder.Writeln("Paragraph after end bookmark.");

        // Save the source document (optional, helps debugging).
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate the start and end bookmarks.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        Bookmark startBookmark = loadedDoc.Range.Bookmarks[startBookmarkName];
        Bookmark endBookmark = loadedDoc.Range.Bookmarks[endBookmarkName];

        if (startBookmark == null)
            throw new InvalidOperationException($"Start bookmark \"{startBookmarkName}\" not found.");
        if (endBookmark == null)
            throw new InvalidOperationException($"End bookmark \"{endBookmarkName}\" not found.");

        // The bookmark start and end nodes are children of paragraphs.
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.ParentNode as Paragraph;

        if (startParagraph == null)
            throw new InvalidOperationException("Start bookmark is not inside a paragraph.");
        if (endParagraph == null)
            throw new InvalidOperationException("End bookmark is not inside a paragraph.");

        // -----------------------------------------------------------------
        // 3. Clone the nodes between the two bookmarks (inclusive) into a new document.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node currentNode = startParagraph;
        while (currentNode != null)
        {
            Node importedNode = importer.ImportNode(currentNode, true);
            resultBody.AppendChild(importedNode);

            if (currentNode == endParagraph)
                break;

            currentNode = currentNode.NextSibling;
        }

        // -----------------------------------------------------------------
        // 4. Save the extracted segment as PDF.
        // -----------------------------------------------------------------
        const string outputPdfPath = "extracted.pdf";
        resultDoc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("Failed to create the extracted PDF file.");
    }
}
