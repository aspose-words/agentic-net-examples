using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 10 rows and 2 columns.
        builder.StartTable();
        for (int i = 1; i <= 10; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 2");

            // End the current row.
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the first (and only) table in the document.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Split the table at row index 5 (the 6th row becomes the first row of the new table).
        int splitRowIndex = 5; // zero‑based index

        // Clone the original table's formatting without its rows.
        Table newTable = (Table)table.Clone(false);

        // Insert the new (empty) table right after the original table.
        table.ParentNode.InsertAfter(newTable, table);

        // Move rows from the split index to the end of the original table into the new table.
        while (table.Rows.Count > splitRowIndex)
        {
            // Remove the row from the original table.
            Row movingRow = table.Rows[splitRowIndex];
            movingRow.Remove();

            // Add the row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Save the resulting document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");
    }
}
