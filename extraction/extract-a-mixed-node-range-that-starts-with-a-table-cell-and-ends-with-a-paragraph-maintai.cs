using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // ---------- Create a sample source document ----------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Intro paragraph
        builder.Writeln("Intro paragraph before the table.");

        // Create a table with two rows and two cells each
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();
        builder.EndTable();

        // Paragraphs after the table
        builder.Writeln("First paragraph after the table.");
        builder.Writeln("Second paragraph after the table.");

        // Save the source document (local file)
        sourceDoc.Save("source.docx");

        // ---------- Load the source document ----------
        Document loadedDoc = new Document("source.docx");

        // Locate the start node: the first table cell
        Cell startCell = loadedDoc.GetChildNodes(NodeType.Cell, true)[0] as Cell;
        if (startCell == null)
            throw new InvalidOperationException("Start cell not found.");

        // The table that contains the start cell will be the first block node in the range
        Table startTable = startCell.ParentNode?.ParentNode as Table;
        if (startTable == null)
            throw new InvalidOperationException("Containing table not found.");

        // Locate the end node: the last paragraph in the document
        Body body = loadedDoc.FirstSection.Body;
        Paragraph endParagraph = body.Paragraphs[body.Paragraphs.Count - 1];
        if (endParagraph == null)
            throw new InvalidOperationException("End paragraph not found.");

        // ---------- Build the extracted document ----------
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Create a new section and body for the result document
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to bring nodes from the source document into the result document
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        Node importedTable = importer.ImportNode(startTable, true);
        Node importedParagraph = importer.ImportNode(endParagraph, true);

        resultBody.AppendChild(importedTable);
        resultBody.AppendChild(importedParagraph);

        // Save the extracted content
        resultDoc.Save("extracted.docx");

        // Validate that the output file was created
        if (!File.Exists("extracted.docx"))
            throw new InvalidOperationException("The extracted document was not created.");
    }
}
