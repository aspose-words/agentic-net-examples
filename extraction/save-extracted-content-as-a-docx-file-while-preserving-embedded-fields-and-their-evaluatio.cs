using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document containing fields and bookmarks.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Insert a DATE field and a line break.
        builder.InsertField(FieldType.FieldDate, true);
        builder.Writeln();

        // Insert a TIME field and a line break.
        builder.InsertField(FieldType.FieldTime, true);
        builder.Writeln();

        // First bookmark – the content we want to extract starts here.
        builder.StartBookmark("ExtractStart");
        builder.Writeln("This paragraph is inside the extraction range.");
        builder.EndBookmark("ExtractStart");

        // Insert a PAGE field that also belongs to the extraction range.
        builder.InsertField(FieldType.FieldPage, true);
        builder.Writeln();

        // Second bookmark – marks the end of the extraction range.
        builder.StartBookmark("ExtractEnd");
        builder.Writeln("End of extraction range.");
        builder.EndBookmark("ExtractEnd");

        // Save the source document to a local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // ------------------------------------------------------------
        // 2. Load the source document and update all fields so that
        //    their displayed results are current.
        // ------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        loadedDoc.UpdateFields();

        // ------------------------------------------------------------
        // 3. Locate the start and end bookmarks that define the range.
        // ------------------------------------------------------------
        Bookmark startBookmark = loadedDoc.Range.Bookmarks["ExtractStart"];
        Bookmark endBookmark = loadedDoc.Range.Bookmarks["ExtractEnd"];
        if (startBookmark == null || endBookmark == null)
            throw new InvalidOperationException("Required bookmarks were not found.");

        // The bookmarks are markers; the actual content we need is the
        // paragraphs (or other block nodes) that are children of the
        // paragraphs containing the bookmark markers.
        Paragraph startParagraph = startBookmark.BookmarkStart.ParentNode as Paragraph;
        Paragraph endParagraph = endBookmark.BookmarkEnd.ParentNode as Paragraph;

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Unable to locate paragraph boundaries.");

        // ------------------------------------------------------------
        // 4. Build a new empty document that will hold the extracted content.
        // ------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Ensure a clean document.

        // Create the minimal required structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // ------------------------------------------------------------
        // 5. Walk through the block-level nodes from the start paragraph
        //    to the end paragraph (inclusive) and import them into the
        //    new document.
        // ------------------------------------------------------------
        Node currentNode = startParagraph;
        while (currentNode != null)
        {
            // Only block-level nodes (Paragraph, Table, etc.) can be added to Body.
            if (currentNode.NodeType == NodeType.Paragraph ||
                currentNode.NodeType == NodeType.Table)
            {
                Node importedNode = extractedDoc.ImportNode(currentNode, true);
                body.AppendChild(importedNode);
            }

            if (currentNode == endParagraph)
                break;

            currentNode = currentNode.NextSibling;
        }

        // ------------------------------------------------------------
        // 6. Save the extracted content as a DOCX file, preserving fields
        //    and their evaluated results.
        // ------------------------------------------------------------
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Verify that the output file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted DOCX file was not created.");

        Console.WriteLine("Extraction completed successfully. Output file: " + extractedPath);
    }
}
