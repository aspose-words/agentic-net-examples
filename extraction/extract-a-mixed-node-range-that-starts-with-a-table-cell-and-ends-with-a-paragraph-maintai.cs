using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractMixedRange
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a sample source document containing a table and paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph.
        builder.Writeln("Intro paragraph.");

        // Table – the first cell will be the start of the extraction range.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("StartCell"); // Start node.
        builder.InsertCell();
        builder.Write("Cell2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell3");
        builder.InsertCell();
        builder.Write("Cell4");
        builder.EndRow();
        builder.EndTable();

        // Paragraphs after the table.
        builder.Writeln("Middle paragraph.");
        builder.Writeln("End paragraph."); // End node.

        // Save the source document to a deterministic local file.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Load the document for extraction.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        // Locate the start cell (first cell of the first table).
        Table firstTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (firstTable == null)
            throw new InvalidOperationException("No table found in the source document.");

        Cell startCell = firstTable.Rows[0].Cells[0];
        if (startCell == null)
            throw new InvalidOperationException("Start cell not found.");

        // Locate the end paragraph (the last paragraph in the document).
        Paragraph endParagraph = loadedDoc.FirstSection.Body.Paragraphs[
            loadedDoc.FirstSection.Body.Paragraphs.Count - 1];
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // Locate the paragraph that follows the table (the middle paragraph).
        Paragraph middleParagraph = firstTable.NextSibling as Paragraph;
        if (middleParagraph == null)
            throw new InvalidOperationException("Middle paragraph not found.");

        // -----------------------------------------------------------------
        // Build the result document that will contain the extracted range.
        // -----------------------------------------------------------------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove the default empty section.

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Import the whole table (contains the start cell) and append it.
        Node importedTable = importer.ImportNode(firstTable, true);
        resultBody.AppendChild(importedTable);

        // Import the middle paragraph and append it.
        Node importedMiddle = importer.ImportNode(middleParagraph, true);
        resultBody.AppendChild(importedMiddle);

        // Import the end paragraph and append it.
        Node importedEnd = importer.ImportNode(endParagraph, true);
        resultBody.AppendChild(importedEnd);

        // Save the extracted content.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
