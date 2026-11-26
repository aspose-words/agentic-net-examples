using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a source document with an outer table that contains
        //    a nested table inside its second cell.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Start outer table.
        builder.StartTable();

        // First cell of outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 1");

        // Second cell of outer table – will contain a nested table.
        builder.InsertCell();

        // Start nested table inside the second cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Nested Cell 1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Nested Cell 2");
        builder.EndRow();
        builder.EndTable(); // End nested table.

        builder.EndRow();   // End outer table row.
        builder.EndTable(); // End outer table.

        // Save the source document.
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // 2. Load the source document and locate the outer table.
        // ------------------------------------------------------------
        Document loadedDoc = new Document("source.docx");
        Table outerTable = loadedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (outerTable == null)
            throw new InvalidOperationException("Outer table not found in the source document.");

        // ------------------------------------------------------------
        // 3. Create a new empty document that will hold the extracted segment.
        // ------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build minimal document structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // ------------------------------------------------------------
        // 4. Import (clone) the outer table into the new document.
        //    Use NodeImporter to transfer the node between documents.
        // ------------------------------------------------------------
        NodeImporter importer = new NodeImporter(loadedDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(outerTable, true);
        if (importedTable == null)
            throw new InvalidOperationException("Failed to import the outer table into the extracted document.");

        body.AppendChild(importedTable);

        // ------------------------------------------------------------
        // 5. Save the extracted document.
        // ------------------------------------------------------------
        extractedDoc.Save("extracted.docx");

        // ------------------------------------------------------------
        // 6. Validation: ensure the extracted document contains the nested table.
        // ------------------------------------------------------------
        Table extractedOuter = extractedDoc.GetChildNodes(NodeType.Table, true)[0] as Table;
        if (extractedOuter == null)
            throw new InvalidOperationException("Extracted outer table not found.");

        // The nested table resides in the second cell of the outer table.
        Cell secondCell = extractedOuter.FirstRow.Cells[1];
        NodeCollection nestedTables = secondCell.GetChildNodes(NodeType.Table, true);
        if (nestedTables == null || nestedTables.Count == 0)
            throw new InvalidOperationException("Nested table was not retained in the extracted document.");

        // ------------------------------------------------------------
        // 7. Verify that both output files were created.
        // ------------------------------------------------------------
        if (!File.Exists("source.docx") || !File.Exists("extracted.docx"))
            throw new InvalidOperationException("Expected output files were not created.");

        // If we reach this point, the extraction succeeded.
        Console.WriteLine("Extraction completed successfully.");
    }
}
