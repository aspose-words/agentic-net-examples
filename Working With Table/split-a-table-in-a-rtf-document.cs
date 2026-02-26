using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInRtf
{
    static void Main()
    {
        // Load the RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Assume we want to split the first table at a specific row index.
        // For example, split after the third row (zero‑based index = 2).
        int splitAfterRowIndex = 2;

        // Get the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Validate the split index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Create a new table that will hold the rows after the split point.
        Table newTable = new Table(doc);
        // Copy basic formatting from the original table to keep appearance consistent.
        newTable.PreferredWidth = originalTable.PreferredWidth;
        newTable.Alignment = originalTable.Alignment;
        newTable.Style = originalTable.Style;
        newTable.StyleIdentifier = originalTable.StyleIdentifier;
        newTable.StyleName = originalTable.StyleName;

        // Move rows from the original table to the new table.
        // Start moving from the row after the split index.
        // While moving, the collection size changes, so always remove the row at splitAfterRowIndex + 1.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // Remove the row from the original table and add it to the new table.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];
            originalTable.Rows.RemoveAt(splitAfterRowIndex + 1);
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table into the document immediately after the original table.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document back to RTF (or any other format you need).
        doc.Save("OutputDocument.rtf");
    }
}
