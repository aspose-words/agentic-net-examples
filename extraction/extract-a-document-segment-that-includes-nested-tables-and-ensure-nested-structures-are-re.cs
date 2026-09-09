using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains nested tables.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Marker paragraph before the tables.
        builder.Writeln("=== Start of Document ===");

        // Build the outer table (2 rows, 2 columns).
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Outer Cell 1");

        // First row, second cell – this cell will contain an inner table.
        builder.InsertCell();

        // Move the cursor into the cell to insert the inner table.
        Cell outerCell = builder.CurrentParagraph.ParentNode as Cell;
        builder.MoveTo(outerCell.FirstParagraph);

        // Build the inner (nested) table (2 rows, 2 columns).
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Inner Cell 1");
        builder.InsertCell();
        builder.Writeln("Inner Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Writeln("Inner Cell 3");
        builder.InsertCell();
        builder.Writeln("Inner Cell 4");
        builder.EndTable();

        // Return to the outer table to finish the first row.
        builder.MoveTo(outerCell.FirstParagraph);
        builder.EndRow();

        // Second row of the outer table.
        builder.InsertCell();
        builder.Writeln("Outer Cell 3");
        builder.InsertCell();
        builder.Writeln("Outer Cell 4");
        builder.EndRow();

        // End the outer table.
        builder.EndTable();

        // Marker paragraph after the tables.
        builder.Writeln("=== End of Document ===");

        // Save the source document.
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the document back.
        Document loadedDoc = new Document(sourcePath);

        // Locate the outer table that contains the nested table.
        Table outerTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table not found in the source document.");

        // Create a new empty document to hold the extracted segment.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build a minimal document structure (Section -> Body).
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Import the outer table (including its nested inner table) into the new document.
        Node importedTable = extractedDoc.ImportNode(outerTable, true);
        body.AppendChild(importedTable);

        // Save the extracted document.
        const string extractedPath = "extracted.docx";
        extractedDoc.Save(extractedPath);

        // Verify that the extracted file was created.
        if (!File.Exists(extractedPath))
            throw new InvalidOperationException("The extracted document was not created.");

        Console.WriteLine("Extraction completed successfully.");
    }
}
