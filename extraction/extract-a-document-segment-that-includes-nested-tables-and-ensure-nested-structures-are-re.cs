using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // -------------------------------------------------
        // Create a source document with a nested table.
        // -------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build the outer 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Outer Cell 1");
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.EndRow();

        // Start the second row (no explicit StartRow method needed).
        builder.InsertCell();
        builder.Write("Outer Cell 3");
        builder.InsertCell();
        builder.Write("Outer Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Insert an inner 2x2 table into the first cell of the outer table.
        Table outerTable = sourceDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table was not created.");

        Cell targetCell = outerTable.FirstRow.FirstCell;
        if (targetCell == null)
            throw new InvalidOperationException("Target cell not found.");

        // Move the cursor to the first paragraph of the target cell.
        builder.MoveTo(targetCell.FirstParagraph);
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Inner Cell 1");
        builder.InsertCell();
        builder.Write("Inner Cell 2");
        builder.EndRow();

        // Second row of the inner table.
        builder.InsertCell();
        builder.Write("Inner Cell 3");
        builder.InsertCell();
        builder.Write("Inner Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // -------------------------------------------------
        // Load the source document and extract the outer table (including its nested table).
        // -------------------------------------------------
        Document loadedDoc = new Document(sourcePath);
        Table tableToExtract = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (tableToExtract == null)
            throw new InvalidOperationException("Table to extract was not found.");

        // Create a new empty document that will hold the extracted segment.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        // Build the minimal required structure: Section -> Body.
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import the table from the source document into the result document.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        Table importedTable = (Table)importer.ImportNode(tableToExtract, true);
        resultBody.AppendChild(importedTable);

        // Save the extracted segment.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // -------------------------------------------------
        // Validation.
        // -------------------------------------------------
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Verify that the nested table is retained.
        int nestedTableCount = importedTable.GetChildNodes(NodeType.Table, true).Count;
        if (nestedTableCount == 0)
            throw new InvalidOperationException("Nested tables were not retained in the extracted document.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
