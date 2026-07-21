using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace SplitTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with five rows and two columns.
            Table table = builder.StartTable();

            for (int row = 1; row <= 5; row++)
            {
                // First cell.
                builder.InsertCell();
                builder.Write($"Row {row}, Cell 1");

                // Second cell.
                builder.InsertCell();
                builder.Write($"Row {row}, Cell 2");

                // End the current row (except after the last row we will end the table later).
                if (row < 5)
                    builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Split the table after the third row (zero‑based index 2).
            int splitIndex = 2; // rows 0,1,2 stay in the original table

            // Create a new table that will hold the rows after the split.
            Table secondTable = new Table(doc);
            // Insert the new table right after the original one in the document tree.
            table.ParentNode.InsertAfter(secondTable, table);

            // Move rows from the original table to the new table.
            while (table.Rows.Count > splitIndex + 1)
            {
                // The row to move is always the one that follows the split index.
                Row rowToMove = table.Rows[splitIndex + 1];
                // Detach the row from the original table.
                rowToMove.Remove();
                // Append it to the new table.
                secondTable.Rows.Add(rowToMove);
            }

            // Validate the split.
            if (table.Rows.Count != 3 || secondTable.Rows.Count != 2)
                throw new InvalidOperationException("Table split did not produce the expected row counts.");

            // Save the document.
            doc.Save("SplitTable.docx");
        }
    }
}
