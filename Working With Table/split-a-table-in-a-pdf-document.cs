using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SplitTableInPdf
{
    static void Main()
    {
        // Load the source PDF (Aspose.Words can load PDF as a document)
        Document doc = new Document("Input.pdf");

        // Assume we want to split the first table in the document after the third row
        const int splitAfterRowIndex = 2; // zero‑based index

        // Get the first table in the first section
        Table originalTable = doc.FirstSection.Body.Tables[0];

        // Validate that the table has enough rows to split
        if (originalTable.Rows.Count <= splitAfterRowIndex + 1)
        {
            Console.WriteLine("The table does not have enough rows to split.");
            return;
        }

        // Create a new table that will hold the rows after the split point.
        // Clone the original table without its child rows (shallow clone) to preserve formatting.
        Table newTable = (Table)originalTable.Clone(false);

        // Move rows from the original table to the new table.
        // Start moving from the row after the split point to the end.
        for (int i = originalTable.Rows.Count - 1; i > splitAfterRowIndex; i--)
        {
            Row rowToMove = originalTable.Rows[i];
            // Remove the row from the original table.
            originalTable.Rows.RemoveAt(i);
            // Append the row to the new table (preserves formatting).
            newTable.Rows.Add(rowToMove);
        }

        // Insert a page break between the two tables.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Move the cursor to the position just after the original table.
        builder.MoveTo(originalTable.LastRow);
        builder.Writeln(); // ensure we are after the table.
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the new table after the page break.
        builder.InsertNode(newTable);

        // Save the modified document as PDF.
        doc.Save("Output.pdf", SaveFormat.Pdf);
    }
}
