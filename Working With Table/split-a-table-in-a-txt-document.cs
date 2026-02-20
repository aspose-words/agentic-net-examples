using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with 4 rows and 2 columns.
        Table table = builder.StartTable();
        for (int i = 1; i <= 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 1");
            builder.InsertCell();
            builder.Write($"Row {i}, Cell 2");
            builder.EndRow();
        }
        builder.EndTable();

        // Split the table after the second row.
        SplitTableAfterRow(table, 2);

        // Save the document as plain text while preserving table layout.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            PreserveTableLayout = true
        };
        doc.Save("SplitTable.txt", saveOptions);
    }

    // Moves rows after splitIndex into a new table inserted after the original table.
    static void SplitTableAfterRow(Table original, int splitIndex)
    {
        // Validate split index.
        if (splitIndex < 1 || splitIndex >= original.Rows.Count)
            return;

        // Clone the original table's formatting without its rows.
        Table newTable = (Table)original.Clone(false);

        // Insert the new table immediately after the original table.
        original.ParentNode.InsertAfter(newTable, original);

        // Number of rows to move to the new table.
        int rowsToMove = original.Rows.Count - splitIndex;

        // Transfer rows one by one.
        for (int i = 0; i < rowsToMove; i++)
        {
            // Always remove the row at the split index (rows shift up after removal).
            Row row = original.Rows[splitIndex];
            original.RemoveChild(row);
            newTable.AppendChild(row);
        }
    }
}
