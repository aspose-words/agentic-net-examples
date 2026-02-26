using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableSplitExample
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Input.docx");

        // Get the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the document.");

        // Define the row index at which to split the table.
        // Rows with index < splitIndex stay in the original table,
        // rows with index >= splitIndex move to the new table.
        int splitIndex = 2; // example: split after the second row (zero‑based)

        if (splitIndex < 0 || splitIndex >= originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitIndex), "Split index is out of range.");

        // Clone the original table structure without its rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the split point to the new table.
        // Rows are removed from the original table as they are added to the new one.
        while (originalTable.Rows.Count > splitIndex)
        {
            Row movingRow = originalTable.Rows[splitIndex];
            movingRow.Remove();               // Detach from original table.
            newTable.Rows.Add(movingRow);      // Append to the new table.
        }

        // Insert the new table immediately after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
