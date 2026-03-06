using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInDotm
{
    static void Main()
    {
        // Load the DOTM template.
        Document doc = new Document("Template.dotm");

        // Find the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the document.");

        // Index of the row after which the table will be split (0‑based).
        // Rows with index < splitRow stay in the original table,
        // rows with index >= splitRow move to the new table.
        int splitRow = 2; // example: split after the second row

        // Ensure the split index is valid.
        if (splitRow <= 0 || splitRow >= originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitRow), "Split row index is out of range.");

        // Create a shallow clone of the original table (no rows, cells, etc.).
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // We start at splitRow because rows before it remain in the original.
        while (originalTable.Rows.Count > splitRow)
        {
            // Remove the row from the original table.
            Row movingRow = originalTable.Rows[splitRow];
            originalTable.Rows.Remove(movingRow);

            // Append the removed row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table immediately after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
