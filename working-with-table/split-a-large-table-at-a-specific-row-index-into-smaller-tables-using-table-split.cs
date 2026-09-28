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

        // Build a sample table with 10 rows and 2 columns.
        builder.StartTable();
        for (int i = 1; i <= 10; i++)
        {
            // First cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i} Cell 1");

            // Second cell of the row.
            builder.InsertCell();
            builder.Writeln($"Row {i} Cell 2");

            // End the current row.
            builder.EndRow();
        }
        // End the table construction.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("Table was not created.");

        // Index at which to split the table (zero‑based). Rows from this index onward will move to a new table.
        int splitIndex = 5;

        // Clone the original table without its rows to obtain a new empty table with the same formatting.
        Table newTable = (Table)originalTable.Clone(false);
        // Insert the new table right after the original table in the document.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Move rows starting from splitIndex from the original table to the new table.
        while (originalTable.Rows.Count > splitIndex)
        {
            Row rowToMove = originalTable.Rows[splitIndex];
            // Remove the row from the original table.
            originalTable.Rows.RemoveAt(splitIndex);
            // Append the removed row to the new table.
            newTable.Rows.Add(rowToMove);
        }

        // Validate row counts of the resulting tables.
        int originalRowCount = originalTable.Rows.Count; // Expected: 5
        int newRowCount = newTable.Rows.Count;           // Expected: 5

        Console.WriteLine($"Original table rows after split: {originalRowCount}");
        Console.WriteLine($"New table rows after split: {newRowCount}");

        // Save the document containing the split tables.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
