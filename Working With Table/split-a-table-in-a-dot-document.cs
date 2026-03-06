using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOC/DOT document.
        Document doc = new Document("InputDocument.dot");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index after which the table will be split.
        // For example, split after the second row (zero‑based index = 1).
        int splitAfterRowIndex = 1;

        // If the table has fewer rows than the split point, nothing to do.
        if (originalTable.Rows.Count <= splitAfterRowIndex + 1)
        {
            doc.Save("OutputDocument.docx");
            return;
        }

        // Create a shallow clone of the original table (no rows, cells, etc.).
        Table newTable = (Table)originalTable.Clone(false);

        // Move all rows that come after the split point to the new table.
        // Rows are zero‑based; the row at splitAfterRowIndex stays in the original table.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // The row that now occupies the position after the split point.
            Row movingRow = originalTable.Rows[splitAfterRowIndex + 1];

            // Detach the row from the original table.
            movingRow.Remove();

            // Append the detached row to the new table.
            newTable.Rows.Add(movingRow);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
