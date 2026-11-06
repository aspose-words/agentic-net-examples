using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a source document with a nested table.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Outer table start.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Outer Cell 1");

        // Second cell will contain the inner table.
        builder.InsertCell();

        // Inner table inside the current cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Inner Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Inner Cell 2");
        builder.EndRow();
        builder.EndTable();

        // End first row of outer table.
        builder.EndRow();

        // Second row of outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.InsertCell();
        builder.Write("Outer Cell 3");
        builder.EndRow();

        // End outer table.
        builder.EndTable();

        // Save the source document.
        string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedDoc = new Document(sourcePath);

        // Retrieve the outer table (first table in the document).
        Table outerTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table not found.");

        // Deep clone the outer table to retain nested structures.
        Table clonedTable = (Table)outerTable.Clone(true);

        // Create a new document to hold the extracted segment.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove default section and body.

        // Build a minimal document structure.
        Section section = new Section(resultDoc);
        resultDoc.AppendChild(section);
        Body body = new Body(resultDoc);
        section.AppendChild(body);

        // Import the cloned table into the new document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(clonedTable, true);
        body.AppendChild(importedTable);

        // Save the extracted document.
        string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the extracted document still contains both tables.
        Document verifyDoc = new Document(resultPath);
        int tableCount = verifyDoc.GetChildNodes(NodeType.Table, true).Count;
        if (tableCount < 2)
            throw new InvalidOperationException("Nested tables were not retained in the extracted document.");
    }
}
