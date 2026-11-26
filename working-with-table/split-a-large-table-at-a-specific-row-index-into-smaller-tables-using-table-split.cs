using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 6 rows and 2 columns.
        builder.StartTable();
        for (int i = 1; i <= 6; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 1");
            builder.InsertCell();
            builder.Writeln($"Row {i}, Cell 2");
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Split the table at row index 3 (zero‑based).
        // Rows 0‑2 remain in the original table, rows 3‑5 move to a new table.
        int splitIndex = 3;
        if (splitIndex < 0 || splitIndex > originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitIndex), "Invalid split index.");

        // Create a new empty table with the same formatting as the original.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Collect rows to move to preserve order.
        var rowsToMove = new Row[originalTable.Rows.Count - splitIndex];
        for (int i = splitIndex; i < originalTable.Rows.Count; i++)
        {
            rowsToMove[i - splitIndex] = originalTable.Rows[i];
        }

        // Remove rows from the original table and add them to the new table.
        foreach (var row in rowsToMove)
        {
            originalTable.Rows.Remove(row);
            newTable.Rows.Add(row);
        }

        // Insert the new table after the original table in the document.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output file was not created.");

        // Output row counts for verification.
        Console.WriteLine($"Original table rows after split: {originalTable.Rows.Count}");
        Console.WriteLine($"New table rows after split: {newTable.Rows.Count}");
    }
}
