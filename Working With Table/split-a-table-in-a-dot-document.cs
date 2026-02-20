using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableExample
{
    static void Main()
    {
        // Load the DOT (Word template) document.
        Document doc = new Document("Template.dot");

        // Assume we want to split the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split the table.
        // Rows with index < splitIndex stay in the original table,
        // rows with index >= splitIndex move to the new table.
        int splitIndex = 2; // example: split after the second row

        // Guard against invalid split index.
        if (splitIndex <= 0 || splitIndex >= originalTable.Rows.Count)
        {
            Console.WriteLine("Invalid split index.");
            return;
        }

        // Create a new empty table that copies the formatting of the original table.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Iterate from the end to avoid index shifting when removing rows.
        for (int i = originalTable.Rows.Count - 1; i >= splitIndex; i--)
        {
            Row row = originalTable.Rows[i];
            // Remove the row from the original table.
            originalTable.Rows.RemoveAt(i);
            // Insert the row at the beginning of the new table to preserve order.
            newTable.Rows.Insert(0, row);
        }

        // Insert the new table into the document immediately after the original table.
        Node nextNode = originalTable.NextSibling;
        if (nextNode != null)
            originalTable.ParentNode.InsertAfter(newTable, originalTable);
        else
            originalTable.ParentNode.AppendChild(newTable);

        // Save the modified document.
        doc.Save("Template_SplitTable.dot");
    }
}
