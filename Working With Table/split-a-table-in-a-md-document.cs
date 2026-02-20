using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SplitTableInMarkdown
{
    static void Main()
    {
        // Load the Markdown document.
        Document doc = new Document("Input.md");

        // Get the first table in the document.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define the row index after which the table will be split.
        // Rows are zero‑based; this example splits after the third row (index 2).
        int splitAfterRowIndex = 2;

        // Create a new empty table that will hold the rows after the split point.
        Table newTable = (Table)originalTable.Clone(false); // clone without child rows

        // Move rows from the original table to the new table, starting after the split index.
        // Iterate backwards to avoid index shifting while removing rows.
        for (int i = originalTable.Rows.Count - 1; i > splitAfterRowIndex; i--)
        {
            Row row = originalTable.Rows[i];
            originalTable.Rows.RemoveAt(i);
            // Insert at the beginning of the new table to preserve original order.
            newTable.Rows.Insert(0, row);
        }

        // Insert the new table into the document immediately after the original table.
        // ParentNode of a Table is a CompositeNode (e.g., Body), which provides InsertAfter.
        CompositeNode parent = (CompositeNode)originalTable.ParentNode;
        parent.InsertAfter(newTable, originalTable);

        // Save the modified document back to Markdown format.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("Output.md", saveOptions);
    }
}
