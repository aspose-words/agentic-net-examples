using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOCX document.
        Document doc = new Document("Input.docx");

        // Locate the first table in the document.
        Table originalTable = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (originalTable == null)
            throw new InvalidOperationException("No table found in the document.");

        // Define the row index at which to split the table.
        // Rows with index < splitIndex will stay in the original table,
        // rows with index >= splitIndex will move to the new table.
        int splitIndex = 2; // Example: split after the first two rows.

        // Clone the original table to create a new table that will hold the second part.
        Table newTable = (Table)originalTable.Clone(true);

        // Remove rows from the new table that belong to the first part.
        for (int i = 0; i < splitIndex && newTable.Rows.Count > 0; i++)
            newTable.Rows.RemoveAt(0);

        // Remove rows from the original table that belong to the second part.
        for (int i = originalTable.Rows.Count - 1; i >= splitIndex; i--)
            originalTable.Rows.RemoveAt(i);

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
