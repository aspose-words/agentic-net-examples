using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class SplitTableInMarkdown
{
    static void Main()
    {
        // Load the Markdown document.
        Document doc = new Document("input.md");

        // Assume the document contains at least one table.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index at which to split the table.
        // Rows before this index stay in the original table,
        // rows from this index onward move to a new table.
        int splitRowIndex = 2; // Example: split after the first two rows.

        // Create a new table that will hold the rows after the split point.
        Table newTable = new Table(doc);

        // Collect rows to move to the new table.
        List<Row> rowsToMove = new List<Row>();
        for (int i = splitRowIndex; i < originalTable.Rows.Count; i++)
        {
            rowsToMove.Add(originalTable.Rows[i]);
        }

        // Move each collected row from the original table to the new table.
        foreach (Row row in rowsToMove)
        {
            // Remove the row from the original table.
            originalTable.RemoveChild(row);
            // Append the row to the new table.
            newTable.AppendChild(row);
        }

        // Insert the new table immediately after the original table in the document.
        // The InsertAfter method belongs to CompositeNode, so cast the parent accordingly.
        CompositeNode parent = originalTable.ParentNode as CompositeNode;
        if (parent != null)
        {
            parent.InsertAfter(newTable, originalTable);
        }
        else
        {
            // Fallback: add the new table to the document body if casting fails.
            doc.FirstSection.Body.AppendChild(newTable);
        }

        // Save the modified document back to Markdown format.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("output.md", saveOptions);
    }
}
