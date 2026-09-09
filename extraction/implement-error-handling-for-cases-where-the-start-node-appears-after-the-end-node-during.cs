using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with two bookmarks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Paragraph 1");
        builder.StartBookmark("Start");
        builder.Writeln("Paragraph 2");
        builder.EndBookmark("Start");
        builder.StartBookmark("End");
        builder.Writeln("Paragraph 3");
        builder.EndBookmark("End");

        string sourcePath = "sample.docx";
        doc.Save(sourcePath);

        // Load the document for extraction.
        Document loaded = new Document(sourcePath);

        // Retrieve the bookmarks that define the extraction boundaries.
        Bookmark startBookmark = loaded.Range.Bookmarks["Start"];
        Bookmark endBookmark = loaded.Range.Bookmarks["End"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks not found.");

        // Determine the paragraphs that contain the bookmark starts.
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkStart.ParentNode as Paragraph;
        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Bookmarks are not located inside paragraphs.");

        // Find the positions of the start and end paragraphs within the body.
        Body body = loaded.FirstSection.Body;
        NodeCollection paragraphs = body.GetChildNodes(NodeType.Paragraph, true);
        int startIndex = paragraphs.IndexOf(startParagraph);
        int endIndex = paragraphs.IndexOf(endParagraph);

        // Validate ordering: start must precede end.
        if (startIndex > endIndex)
        {
            Console.WriteLine("Error: The start node appears after the end node. Extraction aborted.");
            return;
        }

        // Build a new document containing the extracted range.
        Document result = new Document();
        result.RemoveAllChildren();

        Section resultSection = new Section(result);
        result.AppendChild(resultSection);

        Body resultBody = new Body(result);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loaded, result, ImportFormatMode.KeepSourceFormatting);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = (Paragraph)paragraphs[i];
            Node importedNode = importer.ImportNode(srcParagraph, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted content.
        string outputPath = "extracted.docx";
        result.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Extraction output file was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
