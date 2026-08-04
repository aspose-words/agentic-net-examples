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

        // Build a table with 5 rows and 2 columns.
        Table table = builder.StartTable();
        for (int i = 1; i <= 5; i++)
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

        // Split the table at the third row (zero‑based index 2).
        // The original table will keep rows before the split,
        // and the new table will contain the rows from the split index onward.
        int splitIndex = 2; // zero‑based index where the split starts

        // Create a new empty table that will receive the split rows.
        Table secondTable = new Table(doc);

        // Move rows from the original table to the new table.
        while (table.Rows.Count > splitIndex)
        {
            // Get the row that should be moved.
            Row movingRow = table.Rows[splitIndex];

            // Detach the row from the original table.
            movingRow.Remove();

            // Append the row to the new table.
            secondTable.Rows.Add(movingRow);
        }

        // Insert the new table into the document immediately after the original table.
        table.ParentNode.InsertAfter(secondTable, table);

        // Optional validation: output row counts of both tables.
        Console.WriteLine($"First table rows: {table.Rows.Count}");
        Console.WriteLine($"Second table rows: {secondTable.Rows.Count}");

        // Save the document containing the two separate tables.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
