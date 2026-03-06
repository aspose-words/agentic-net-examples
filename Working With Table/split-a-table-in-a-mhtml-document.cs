using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the source MHTML document.
        Document doc = new Document("Input.mhtml");

        // Split the first table after the second row (zero‑based index).
        SplitTable(doc, tableIndex: 0, splitRowIndex: 2);

        // Save the modified document back to MHTML.
        doc.Save("Output.mhtml");
    }

    /// <summary>
    /// Splits a table into two tables at the specified row index.
    /// Rows from <paramref name="splitRowIndex"/> to the end are moved to a new table,
    /// which is inserted immediately after the original table.
    /// </summary>
    /// <param name="doc">The document containing the table.</param>
    /// <param name="tableIndex">Zero‑based index of the table to split within the document body.</param>
    /// <param name="splitRowIndex">Zero‑based index of the row at which to split. This row becomes the first row of the new table.</param>
    static void SplitTable(Document doc, int tableIndex, int splitRowIndex)
    {
        // Retrieve the target table.
        Table original = doc.FirstSection.Body.Tables[tableIndex];

        // Validate the split index.
        if (splitRowIndex < 0 || splitRowIndex >= original.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitRowIndex));

        // Create a new table that will hold the rows after the split point.
        Table newTable = new Table(doc);

        // Preserve style information (optional but keeps appearance consistent).
        newTable.Style = original.Style;
        newTable.StyleIdentifier = original.StyleIdentifier;
        newTable.StyleName = original.StyleName;

        // Move rows from splitRowIndex to the end of the original table into the new table.
        // We remove rows from the end to maintain the original order when inserting at the beginning of the new table.
        while (original.Rows.Count > splitRowIndex)
        {
            Row row = original.LastRow;
            original.RemoveChild(row);
            newTable.PrependChild(row);
        }

        // Insert the new table directly after the original table in the document tree.
        // InsertAfter is defined on CompositeNode, not on the base Node class.
        CompositeNode parent = original.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("The table's parent node is not a CompositeNode.");
        parent.InsertAfter(newTable, original);
    }
}
