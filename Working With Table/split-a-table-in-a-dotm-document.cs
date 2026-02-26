using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplit
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template.
            Document doc = new Document(@"C:\Docs\Template.dotm");

            // Find the first table in the document.
            Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (originalTable == null)
                throw new InvalidOperationException("No table found in the document.");

            // Define the row index at which to split the table.
            // Rows are zero‑based; this example splits after the second row (index 1).
            int splitAfterRowIndex = 1;

            // Ensure the split index is valid.
            if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
                throw new ArgumentOutOfRangeException(nameof(splitAfterRowIndex), "Split index is out of range.");

            // Create a new empty table that will hold the rows after the split point.
            // Clone the original table without its child nodes (rows) to preserve formatting.
            Table newTable = (Table)originalTable.Clone(false);

            // Move rows from the original table to the new table.
            // Start moving from the row after the split index.
            while (originalTable.Rows.Count > splitAfterRowIndex + 1)
            {
                // Remove the row from the original table.
                Row movingRow = originalTable.Rows[splitAfterRowIndex + 1];
                originalTable.Rows.Remove(movingRow);

                // Append the removed row to the new table.
                newTable.Rows.Add(movingRow);
            }

            // Insert the new table immediately after the original table in the document tree.
            originalTable.ParentNode.InsertAfter(newTable, originalTable);

            // Save the modified document. The output format can be DOCX, DOCM, etc.
            doc.Save(@"C:\Docs\Template_Split.docx");
        }
    }
}
