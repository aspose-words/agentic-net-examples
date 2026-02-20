using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Assume the first table in the first section is the one we want to split.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split the table (0‑based).
        // Rows before this index stay in the original table,
        // rows from this index onward move to a new table.
        int splitRowIndex = 2; // example: split after the second row

        // Guard against invalid split positions.
        if (splitRowIndex <= 0 || splitRowIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Clone the table structure without any rows.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // After a row is added to the new table it is automatically removed from the original one.
        while (originalTable.Rows.Count > splitRowIndex)
        {
            Row rowToMove = originalTable.Rows[splitRowIndex];
            newTable.Rows.Add(rowToMove);
        }

        // Insert the new table after the original table.
        // InsertAfter is defined on CompositeNode, so cast the parent accordingly.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent != null)
        {
            parent.InsertAfter(newTable, originalTable);
        }
        else
        {
            Console.WriteLine("Unable to insert the new table – parent node is not a CompositeNode.");
            return;
        }

        // Save the modified document.
        doc.Save("SplitTable.docx");
    }
}
