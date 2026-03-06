using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplit
{
    class Program
    {
        static void Main()
        {
            // Load the HTML document that contains the table to be split.
            Document doc = new Document("Input.html");

            // Find the first table in the document (adjust the index if you need a different table).
            Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (originalTable == null)
                throw new InvalidOperationException("No table found in the document.");

            // Define after which row the table should be split.
            // For example, split after the second row (zero‑based index = 1).
            int splitAfterRowIndex = 1;

            // Validate the split index.
            if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
                throw new ArgumentOutOfRangeException(nameof(splitAfterRowIndex), "Split index is out of range.");

            // Create a new empty table that will receive the rows after the split point.
            // Clone the original table's formatting but not its child nodes.
            Table newTable = (Table)originalTable.Clone(false);

            // Move rows from the original table to the new table, preserving order.
            // Rows after the split index are removed from the original and appended to the new table.
            while (originalTable.Rows.Count > splitAfterRowIndex + 1)
            {
                // The row that follows the split point is always at index splitAfterRowIndex + 1.
                Row movingRow = originalTable.Rows[splitAfterRowIndex + 1];
                movingRow.Remove();               // Detach from the original table.
                newTable.AppendChild(movingRow);   // Append to the new table.
            }

            // Insert the new table directly after the original table in the document tree.
            originalTable.ParentNode.InsertAfter(newTable, originalTable);

            // Save the modified document back to HTML.
            // HtmlSaveOptions can be used if you need to control the output format further.
            doc.Save("Output.html", SaveFormat.Html);
        }
    }
}
