using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableInMarkdown
{
    static void Main()
    {
        // Path to the source Markdown document that contains a table.
        const string inputPath = @"C:\Docs\Input.md";

        // Path where the resulting Markdown document will be saved.
        const string outputPath = @"C:\Docs\Output.md";

        // Load the Markdown document. Aspose.Words can directly load .md files.
        Document doc = new Document(inputPath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count == 0)
        {
            Console.WriteLine("No tables found in the document.");
            return;
        }

        // Get the first table to split. Adjust the index if you need a different table.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define after which row the table should be split.
        // For example, split after the second row (zero‑based index = 1).
        int splitAfterRowIndex = 1;

        // Validate the split index.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Invalid split index. The table must have at least two rows after the split point.");
            return;
        }

        // Create a new table that will hold the rows after the split point.
        Table newTable = new Table(doc);

        // Copy the formatting of the original table to the new one (optional).
        newTable.Style = originalTable.Style;
        newTable.Alignment = originalTable.Alignment;
        newTable.PreferredWidth = originalTable.PreferredWidth;
        newTable.AllowAutoFit = originalTable.AllowAutoFit;

        // Move rows from the original table to the new table.
        // Rows are moved starting from splitAfterRowIndex + 1 because the row at splitAfterRowIndex
        // stays in the original table.
        while (originalTable.Rows.Count > splitAfterRowIndex + 1)
        {
            // The row to move is always at position splitAfterRowIndex + 1 after each removal.
            Row rowToMove = originalTable.Rows[splitAfterRowIndex + 1];

            // Detach the row from the original table.
            rowToMove.Remove();

            // Append the detached row to the new table.
            newTable.AppendChild(rowToMove);
        }

        // Insert the new table immediately after the original table in the document tree.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document back to Markdown format.
        doc.Save(outputPath);
    }
}
