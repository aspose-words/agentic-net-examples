using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SplitTableAndSavePdf
{
    static void Main()
    {
        // Load the source document (Word format) that contains the table to split.
        Document doc = new Document("Input.docx");

        // Assume we want to split the first table in the first section.
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Define after which row the split should occur.
        // For example, split after the third row (zero‑based index 2).
        int splitAfterRowIndex = 2;

        // Guard against invalid split positions.
        if (splitAfterRowIndex < 0 || splitAfterRowIndex >= originalTable.Rows.Count - 1)
        {
            Console.WriteLine("Invalid split position.");
            return;
        }

        // Create a new table that will hold the rows after the split point.
        Table newTable = (Table)originalTable.Clone(false); // Clone only the table structure, not its rows.

        // Move rows from the original table to the new table.
        // Rows are moved starting from splitAfterRowIndex + 1 to the end.
        for (int i = originalTable.Rows.Count - 1; i > splitAfterRowIndex; i--)
        {
            Row rowToMove = originalTable.Rows[i];
            originalTable.Rows.RemoveAt(i);
            newTable.Rows.Insert(0, rowToMove);
        }

        // Insert the new table into the document immediately after the original table.
        // The original table's parent is a Body node, so we can use InsertAfter.
        originalTable.ParentNode.InsertAfter(newTable, originalTable);

        // Save the modified document as a PDF.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
