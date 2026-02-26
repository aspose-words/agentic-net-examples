using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableSplitExample
{
    class Program
    {
        static void Main()
        {
            // Load the WORDML (WordprocessingML) document.
            Document doc = new Document("InputDocument.xml");

            // Find the first table in the document.
            Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (originalTable == null)
                throw new InvalidOperationException("No table found in the document.");

            // Define the row index at which to split the table.
            // Rows with index < splitRowIndex stay in the original table,
            // rows with index >= splitRowIndex move to the new table.
            int splitRowIndex = 2; // Example: split after the second row (0‑based index).

            // Ensure the split index is within the valid range.
            if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
                throw new ArgumentOutOfRangeException(nameof(splitRowIndex), "Split index must be between 1 and the number of rows - 1.");

            // Create a new empty table that will hold the rows after the split.
            // Clone the original table without its child nodes (rows) to preserve formatting.
            Table newTable = (Table)originalTable.Clone(false);

            // Move rows from the original table to the new table, starting from the split index.
            // Iterate backwards to avoid index shifting when removing rows.
            for (int i = originalTable.Rows.Count - 1; i >= splitRowIndex; i--)
            {
                Row rowToMove = originalTable.Rows[i];
                rowToMove.Remove();               // Detach the row from the original table.
                newTable.Rows.Insert(0, rowToMove); // Insert at the beginning to preserve original order.
            }

            // Insert the new table immediately after the original table in the document tree.
            originalTable.ParentNode.InsertAfter(newTable, originalTable);

            // Save the modified document. The format can be any supported by Aspose.Words.
            doc.Save("OutputDocument.docx");
        }
    }
}
