using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInRtf
{
    static void Main()
    {
        // Load the RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Index of the row where the split should occur (zero‑based).
        // Rows with index < splitRowIndex stay in the original table,
        // rows with index >= splitRowIndex move to the new table.
        int splitRowIndex = 2;

        // Clone the original table including its rows.
        Table newTable = (Table)originalTable.Clone(true);

        // Remove rows that should stay in the original table from the cloned copy.
        for (int i = 0; i < splitRowIndex; i++)
        {
            // The first row of the cloned table corresponds to the row we want to discard.
            newTable.FirstRow.Remove();
        }

        // Remove rows that belong to the new table from the original table.
        while (originalTable.Rows.Count > splitRowIndex)
        {
            // The last row of the original table is the one that should be moved.
            originalTable.LastRow.Remove();
        }

        // Insert the new table immediately after the original one.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document back to RTF (or any other format you need).
        doc.Save("OutputDocument.rtf");
    }
}
