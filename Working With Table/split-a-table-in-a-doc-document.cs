using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the zero‑based index of the row after which the table will be split.
        // For example, split after the third row (index 2) – rows 0‑2 stay in the original table,
        // rows 3‑end will be moved to a new table.
        int splitAfterRowIndex = 2;

        // Guard against invalid split positions.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
            throw new ArgumentOutOfRangeException(nameof(splitAfterRowIndex), "Split index is out of range.");

        // Create a new empty table that will receive the rows after the split point.
        Table newTable = new Table(doc);

        // Copy visual formatting from the original table to the new one (optional).
        newTable.Style = originalTable.Style;
        newTable.StyleIdentifier = originalTable.StyleIdentifier;
        newTable.StyleName = originalTable.StyleName;
        newTable.AllowAutoFit = originalTable.AllowAutoFit;
        newTable.Alignment = originalTable.Alignment;
        newTable.PreferredWidth = originalTable.PreferredWidth;
        newTable.CellSpacing = originalTable.CellSpacing;
        newTable.Bidi = originalTable.Bidi;
        newTable.Title = originalTable.Title;
        newTable.Description = originalTable.Description;

        // Move rows that belong to the new table.
        // Rows after the split index are removed from the original table and added to the new table.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // The row to move is always the one that follows the split index,
            // because rows shift left as we remove them.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];
            originalTable.Rows.Remove(rowToMove);
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.doc");
    }
}
