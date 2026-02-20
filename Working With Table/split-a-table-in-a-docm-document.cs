using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInDocm
{
    static void Main()
    {
        // Load the DOCM document.
        Document doc = new Document(@"C:\Docs\InputDocument.docm");

        // Get the first table in the document (adjust the index if needed).
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Row index where the split should occur (zero‑based).
        // All rows with an index >= splitRowIndex will be moved to a new table.
        int splitRowIndex = 3;

        // Basic validation of the split index.
        if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split index – the table cannot be split at the specified row.");
            return;
        }

        // Create a new empty table that copies the formatting of the original table.
        // Clone(false) copies the table's properties but does NOT copy its rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move the rows that belong to the second part of the split into the new table.
        // We iterate backwards so that removing rows does not affect the loop index.
        for (int i = originalTable.Rows.Count - 1; i >= splitRowIndex; i--)
        {
            Row row = originalTable.Rows[i];
            originalTable.Rows.RemoveAt(i);
            // Insert at the beginning of the new table to preserve the original order.
            newTable.Rows.Insert(0, row);
        }

        // Insert the newly created table directly after the original one in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document back to DOCM format.
        doc.Save(@"C:\Docs\OutputDocument.docm");
    }
}
