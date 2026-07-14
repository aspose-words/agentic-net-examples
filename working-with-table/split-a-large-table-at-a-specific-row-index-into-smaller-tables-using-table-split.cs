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
            builder.InsertCell();
            builder.Writeln($"Row {i} - Cell 1");
            builder.InsertCell();
            builder.Writeln($"Row {i} - Cell 2");
            builder.EndRow();
        }
        builder.EndTable();

        // Retrieve the created table.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        int originalRowCount = originalTable.Rows.Count;

        // Split the table at row index 5 (0‑based). Rows 0‑4 stay in the original table,
        // rows 5‑9 move to the new table.
        int splitIndex = 5;
        if (splitIndex < 0 || splitIndex > originalRowCount)
            throw new ArgumentOutOfRangeException(nameof(splitIndex), "Split index is out of range.");

        // Create a new empty table that copies the formatting of the original table.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows starting from splitIndex to the new table.
        // Iterate from the end to avoid index shifting when removing rows.
        for (int i = originalRowCount - 1; i >= splitIndex; i--)
        {
            Row row = originalTable.Rows[i];
            row.Remove();               // Detach from original table.
            newTable.Rows.Add(row);      // Append to the new table.
        }

        // Insert the new table into the document immediately after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Validate the split.
        int firstPartRows = originalTable.Rows.Count;
        int secondPartRows = newTable.Rows.Count;

        if (firstPartRows != splitIndex)
            throw new InvalidOperationException($"Expected {splitIndex} rows in the first part, but found {firstPartRows}.");

        if (secondPartRows != originalRowCount - splitIndex)
            throw new InvalidOperationException($"Expected {originalRowCount - splitIndex} rows in the second part, but found {secondPartRows}.");

        // Save the document.
        string outputPath = "SplitTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
