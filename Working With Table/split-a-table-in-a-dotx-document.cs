using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("input.dotx");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split (e.g., after the second row).
        int splitAfterRowIndex = 1; // zero‑based index

        // Create a new table that will hold the rows after the split point.
        // Clone the original table's formatting but not its child rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Start moving from the row after the split point until no rows remain.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // Remove the row from the original table and insert it into the new table.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];
            originalTable.Rows.RemoveAt(splitAfterRowIndex + 1);
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table into the document immediately after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
