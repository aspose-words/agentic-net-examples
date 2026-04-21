using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with 10 rows and 2 columns.
        Table table = builder.StartTable();
        for (int i = 1; i <= 10; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");

            // End the current row.
            builder.EndRow();
        }
        // Finish the table.
        builder.EndTable();

        // Split the table after the 4th row (zero‑based index).
        int splitRowIndex = 4; // rows 0‑3 stay in the original table, rows 4‑9 move to a new table.

        // Create a new table that will hold the split rows.
        Table newTable = new Table(doc);
        // Insert the new table right after the original one in the document tree.
        table.ParentNode.InsertAfter(newTable, table);

        // Move rows from the original table to the new table.
        // While the original table has more rows than the split index, keep moving the row at splitRowIndex.
        while (table.Rows.Count > splitRowIndex)
        {
            // Get the row that should be moved.
            Row rowToMove = table.Rows[splitRowIndex];
            // Detach the row from the original table.
            rowToMove.Remove();
            // Append the row to the new table.
            newTable.Rows.Add(rowToMove);
        }

        // Optional: verify the split by writing row counts to the console.
        Console.WriteLine($"Original table rows after split: {table.Rows.Count}");
        Console.WriteLine($"New table rows after split: {newTable.Rows.Count}");

        // Save the document containing the two resulting tables.
        doc.Save("SplitTable.docx");
    }
}
