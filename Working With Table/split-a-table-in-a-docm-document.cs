using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOCM document (uses the provided load rule).
        Document doc = new Document("Input.docm");

        // Find the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Define the row index after which the table will be split.
        // For example, split after the second row (index 2, zero‑based).
        int splitRowIndex = 2;
        if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Create a new empty table that will hold the rows after the split.
        // Clone the original table without its child nodes (rows) to preserve formatting.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Continue moving the row at the split index until the original table has only the rows before the split.
        while (originalTable.Rows.Count > splitRowIndex)
        {
            // Remove the row from the original table.
            Row movingRow = originalTable.Rows[splitRowIndex];
            originalTable.Rows.RemoveAt(splitRowIndex);

            // Append the removed row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document (uses the provided save rule).
        doc.Save("Output.docm");
    }
}
