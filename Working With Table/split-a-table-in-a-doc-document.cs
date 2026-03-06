using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("input.docx");

        // Ensure the document contains at least one table.
        if (doc.FirstSection.Body.Tables.Count == 0)
            throw new InvalidOperationException("The document does not contain any tables.");

        // Get the first table to split.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split the table.
        // Rows with index < splitRowIndex will stay in the original table,
        // rows with index >= splitRowIndex will move to the new table.
        int splitRowIndex = 2; // Example: split after the second row (zero‑based).

        if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
            throw new ArgumentOutOfRangeException(nameof(splitRowIndex), "Split index must be within the table rows range.");

        // Clone the original table to create a new table that will hold the second part.
        Table newTable = (Table)originalTable.Clone(true);

        // Remove rows that belong to the second part from the original table.
        // Iterate backwards to avoid index shifting.
        for (int i = originalTable.Rows.Count - 1; i >= splitRowIndex; i--)
        {
            originalTable.Rows[i].Remove();
        }

        // Remove rows that belong to the first part from the new table.
        for (int i = splitRowIndex - 1; i >= 0; i--)
        {
            newTable.Rows[i].Remove();
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
