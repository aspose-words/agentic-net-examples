using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInDotx
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Find the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
        {
            Console.WriteLine("No table found in the document.");
            return;
        }

        // Index after which the table should be split (zero‑based).
        int splitAfterRowIndex = 2; // rows 0,1,2 stay in the original table.

        // Guard against invalid index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Split index is out of range.");
            return;
        }

        // Create a new empty table that copies the original table's formatting.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows that belong to the new table.
        // Iterate backwards so that removal does not affect the loop.
        for (int i = originalTable.Rows.Count - 1; i > splitAfterRowIndex; i--)
        {
            Row row = originalTable.Rows[i];
            // Remove the row from the original table.
            originalTable.Rows.RemoveAt(i);
            // Append the row to the new table.
            newTable.Rows.Add(row);
        }

        // Insert the new table immediately after the original one.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document (DOTX or DOCX).
        doc.Save("Template_Split.dotx");
    }
}
