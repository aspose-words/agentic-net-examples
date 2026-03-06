using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableSplitter
{
    static void Main()
    {
        // Load the source document (WORDML or any supported format)
        Document doc = new Document("Input.docx");

        // Find the first table in the document
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Index of the row after which the table will be split (0‑based).
        // For example, split after the second row -> splitIndex = 2.
        int splitIndex = 2;

        // Validate split index
        if (splitIndex <= 0 || splitIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Create a new empty table that will hold the rows after the split point.
        // Clone the original table without its child nodes to preserve formatting.
        Table newTable = (Table)originalTable.Clone(false);
        // Ensure the new table has no rows (Clone(false) already does this).
        // Move rows from the original table to the new table.
        for (int i = originalTable.Rows.Count - 1; i >= splitIndex; i--)
        {
            Row row = originalTable.Rows[i];
            // Detach the row from the original table.
            row.Remove();
            // Append the row to the new table, preserving its order.
            newTable.Rows.Insert(0, row);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
