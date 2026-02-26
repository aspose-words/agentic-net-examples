using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableSplitter
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document(@"C:\Input\Sample.docm");

        // Ensure the document contains at least one table.
        if (doc.FirstSection.Body.Tables.Count == 0)
        {
            Console.WriteLine("No tables found in the document.");
            return;
        }

        // Get the first table to split.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define after which row the split should occur (0‑based index).
        // For example, split after the second row (i.e., rows 0 and 1 stay in the original table).
        int splitAfterRowIndex = 1;

        // Validate the split index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Clone the original table's formatting but not its child rows.
        // Clone(false) creates a shallow copy – the new table has the same properties but no rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Move rows that belong to the new table.
        // Rows are moved from the bottom up to preserve their original order.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // Take the last row from the original table.
            Row rowToMove = originalTable.LastRow;

            // Detach the row from the original table.
            originalTable.RemoveChild(rowToMove);

            // Insert the row at the beginning of the new table so that the order remains correct.
            newTable.PrependChild(rowToMove);
        }

        // Save the modified document.
        doc.Save(@"C:\Output\Sample_Split.docm");
    }
}
