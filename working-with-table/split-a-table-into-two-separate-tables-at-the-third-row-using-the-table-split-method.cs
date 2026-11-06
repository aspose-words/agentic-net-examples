using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 5 rows and 2 columns.
        builder.StartTable();
        for (int row = 1; row <= 5; row++)
        {
            // First cell.
            builder.InsertCell();
            builder.Write($"Row {row} Cell 1");

            // Second cell.
            builder.InsertCell();
            builder.Write($"Row {row} Cell 2");

            // End the current row.
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table originalTable = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Split the table at the third row (index 2).
        int splitIndex = 2; // zero‑based index where the second table starts

        // Clone the original table structure (without rows) for the second table.
        Table secondTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the second table.
        while (originalTable.Rows.Count > splitIndex)
        {
            Row rowToMove = originalTable.Rows[splitIndex];
            originalTable.Rows.RemoveAt(splitIndex);
            secondTable.Rows.Add(rowToMove);
        }

        // Insert the new table after the original one.
        originalTable.ParentNode.InsertAfter(secondTable, originalTable);

        // Validate the split.
        if (originalTable.Rows.Count != 2)
            throw new InvalidOperationException($"Original table should have 2 rows after split, but has {originalTable.Rows.Count}.");
        if (secondTable.Rows.Count != 3)
            throw new InvalidOperationException($"Second table should have 3 rows after split, but has {secondTable.Rows.Count}.");

        // Save the document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        Console.WriteLine($"Document saved to '{outputPath}'.");
        Console.WriteLine($"Original table rows: {originalTable.Rows.Count}");
        Console.WriteLine($"Second table rows: {secondTable.Rows.Count}");
    }
}
