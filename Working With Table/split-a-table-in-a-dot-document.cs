using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the source document that contains the table to be split.
        Document doc = new Document("Input.docx");

        // Locate the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the document.");

        // Define the row index after which the table will be split.
        // Rows are zero‑based; this example splits after the second row (index 1).
        int splitAfterRowIndex = 1;

        // Validate the split index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
            throw new ArgumentOutOfRangeException(nameof(splitAfterRowIndex), "Invalid split row index.");

        // Clone the original table structure without its child rows.
        Table newTable = (Table)originalTable.Clone(false);
        // Remove any rows that were cloned by default (should be none, but ensure a clean table).
        newTable.RemoveAllChildren();

        // Move rows that belong to the second part into the new table.
        // The rows to move start at splitAfterRowIndex + 1.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // Remove the row from the original table.
            Row movingRow = originalTable.Rows[splitAfterRowIndex + 1];
            movingRow.Remove();

            // Append the removed row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
