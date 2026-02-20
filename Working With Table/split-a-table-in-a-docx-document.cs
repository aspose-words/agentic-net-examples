using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("input.docx");

        // Get the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Index of the row after which the table will be split (zero‑based).
        // For example, split after the second row (index 2 means rows 0,1,2 stay in the original table).
        int splitRowIndex = 2;

        // Clone the original table without its rows to keep the formatting.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        RowCollection originalRows = originalTable.Rows;
        while (originalRows.Count > splitRowIndex)
        {
            // Take the row at the split index.
            Row rowToMove = originalRows[splitRowIndex];

            // Remove it from the original table.
            rowToMove.Remove();

            // Add it to the new table, preserving order.
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table into the document immediately after the original table.
        // The parent of a Table is a CompositeNode (Body, Section, etc.).
        CompositeNode parent = (CompositeNode)originalTable.ParentNode;
        parent.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
