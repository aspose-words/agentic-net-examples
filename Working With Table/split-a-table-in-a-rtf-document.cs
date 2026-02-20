using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SplitTableInRtf
{
    static void Main()
    {
        // Load the source RTF document.
        Document doc = new Document("InputDocument.rtf");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split the table.
        // Rows with index < splitIndex will stay in the original table,
        // rows with index >= splitIndex will be moved to a new table.
        int splitIndex = 2; // example: split after the second row (zero‑based)

        // Guard against invalid split positions.
        if (splitIndex <= 0 || splitIndex >= originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitIndex), "Split index must be within the table rows range.");

        // Create a new empty table that copies the formatting of the original table.
        Table newTable = (Table)originalTable.Clone(false); // shallow clone – copies properties but no rows.

        // Move rows from the original table to the new table.
        // We iterate while the original table still has rows at the split position.
        while (originalTable.Rows.Count > splitIndex)
        {
            // Remove the row from the original table.
            Row movingRow = originalTable.Rows[splitIndex];
            movingRow.Remove();

            // Add the removed row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table into the document immediately after the original table.
        // ParentNode of a Table is a CompositeNode (Body, Section, etc.).
        CompositeNode parent = (CompositeNode)originalTable.ParentNode;
        parent.InsertAfter(newTable, originalTable);

        // Save the modified document as RTF using RtfSaveOptions (optional customisation).
        RtfSaveOptions saveOptions = new RtfSaveOptions
        {
            // Example option: embed generator name.
            ExportGeneratorName = true
        };
        doc.Save("OutputDocument.rtf", saveOptions);
    }
}
