using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the WORDML (or DOCX) document.
        Document doc = new Document("Input.docx");

        // Get the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Index at which to split the table (0‑based). Rows before this index stay in the original table.
        int splitIndex = 2; // Example: split after the second row.

        // Validate the split index.
        if (splitIndex < 0 || splitIndex >= originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitIndex), "Split index is out of range.");

        // Clone the original table's formatting but without any rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the split point to the end of the original table into the new table.
        // Iterate backwards to avoid index shifting while removing rows.
        for (int i = originalTable.Rows.Count - 1; i >= splitIndex; i--)
        {
            Row row = originalTable.Rows[i];
            originalTable.Rows.RemoveAt(i);
            // Insert at the beginning of the new table to preserve original order.
            newTable.Rows.Insert(0, row);
        }

        // Insert the new table directly after the original one in the document tree.
        // The InsertAfter method belongs to CompositeNode, not the base Node class.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent == null)
            throw new InvalidOperationException("The table's parent node is not a CompositeNode.");
        parent.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
