using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a sample table with 6 rows and 3 columns.
        Table originalTable = builder.StartTable();
        for (int row = 0; row < 6; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Split the table after the third row (zero‑based index 3).
        // The original table will keep rows 0‑2, the new table will contain rows 3‑5.
        int splitRowIndex = 3;

        // Create a new empty table that will hold the split rows.
        Table secondTable = new Table(doc);

        // Move rows from the original table to the second table.
        // Continue moving while there are rows at or beyond the split index.
        while (originalTable.Rows.Count > splitRowIndex)
        {
            // Remove the row from the original table.
            Row rowToMove = originalTable.Rows[splitRowIndex];
            originalTable.Rows.Remove(rowToMove);

            // Add the removed row to the new table.
            secondTable.Rows.Add(rowToMove);
        }

        // Insert the new table immediately after the original one in the document body.
        // The parent of a table is a CompositeNode (e.g., Body), so cast to CompositeNode to use InsertAfter.
        ((CompositeNode)originalTable.ParentNode).InsertAfter(secondTable, originalTable);

        // Save the document containing both tables.
        doc.Save("SplitTable.docx");
    }
}
