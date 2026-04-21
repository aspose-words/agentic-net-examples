using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class SplitTableExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 9 rows and 2 columns.
        Table table = builder.StartTable();
        for (int i = 1; i <= 9; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");
            builder.EndRow();
        }
        builder.EndTable();

        // First split: keep first three rows, move the rest to a new table.
        Table secondTable = SplitTable(table, 3);

        // Second split: split the second table after its first three rows.
        Table thirdTable = SplitTable(secondTable, 3);

        // Validate that we have three tables each with three rows.
        if (table.Rows.Count != 3 || secondTable.Rows.Count != 3 || thirdTable.Rows.Count != 3)
            throw new InvalidOperationException("Table split did not produce the expected row counts.");

        // Save the document containing the three split tables.
        doc.Save("SplitTable.docx");

        // Additionally, save each part as a separate document for clarity.
        SaveTablePart(table, "TablePart1.docx");
        SaveTablePart(secondTable, "TablePart2.docx");
        SaveTablePart(thirdTable, "TablePart3.docx");
    }

    // Splits a table by keeping the first 'keepRows' rows in the original table
    // and moving the remaining rows into a newly created table placed after it.
    private static Table SplitTable(Table sourceTable, int keepRows)
    {
        // Clone the source table's formatting (without child rows).
        Table newTable = (Table)sourceTable.Clone(false);

        // Insert the new table directly after the source table in the document tree.
        sourceTable.ParentNode.InsertAfter(newTable, sourceTable);

        // Move rows from the source table to the new table.
        while (sourceTable.Rows.Count > keepRows)
        {
            // Always take the row at the split index; rows shift left after removal.
            Row movingRow = sourceTable.Rows[keepRows];
            sourceTable.Rows.Remove(movingRow);
            newTable.Rows.Add(movingRow);
        }

        return newTable;
    }

    // Helper method to copy a table into a new document and save it.
    private static void SaveTablePart(Table sourceTable, string fileName)
    {
        Document partDoc = new Document();

        // Ensure the document has at least one section.
        if (partDoc.Sections.Count == 0)
            partDoc.AppendChild(new Section(partDoc));

        // Import the table into the new document.
        NodeImporter importer = new NodeImporter(sourceTable.Document, partDoc, ImportFormatMode.KeepSourceFormatting);
        Node importedTable = importer.ImportNode(sourceTable, true);
        partDoc.FirstSection.Body.AppendChild(importedTable);
        partDoc.Save(fileName);
    }
}
