using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Find the first table in the document.
        Table originalTable = doc.GetChild(NodeType.Table, 0, true) as Table;
        if (originalTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Define the row index at which the table will be split.
        // Rows with index >= splitRowIndex will be moved to a new table.
        int splitRowIndex = 2; // example: split after the second row (0‑based)

        // Guard against invalid split positions.
        if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split row index.");
            return;
        }

        // Clone the original table structure (including formatting) but without its rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Iterate backwards to avoid index shifting when removing rows.
        for (int i = originalTable.Rows.Count - 1; i >= splitRowIndex; i--)
        {
            Row row = originalTable.Rows[i];
            row.Remove();               // Detach the row from the original table.
            newTable.Rows.Add(row);     // Append the row to the new table.
        }

        // Insert the new table immediately after the original one.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.doc");
    }
}
