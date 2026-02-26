using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the HTML document that contains a table.
        Document doc = new Document("input.html");

        // Find the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Define after which row the table should be split.
        // For example, split after the first row (index 0).
        int splitAfterRowIndex = 0;

        // Ensure the table has at least one row.
        originalTable.EnsureMinimum();

        // Clone the original table structure without its rows.
        Table newTable = (Table)originalTable.Clone(false);
        // The cloned table has no rows; add a placeholder row to satisfy the minimum requirement.
        newTable.EnsureMinimum();

        // Move rows from the original table to the new table starting after the split index.
        // Rows are moved one by one to preserve formatting and references.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // The row to move is the one immediately after the split point.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];
            // Remove the row from the original table.
            originalTable.Rows.Remove(rowToMove);
            // Add the row to the new table.
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table into the document right after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
