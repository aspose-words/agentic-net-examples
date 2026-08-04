using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableSplitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with 10 rows and 2 columns.
            Table table = builder.StartTable();

            for (int i = 1; i <= 10; i++)
            {
                // First cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Column 1");

                // Second cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Column 2");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table construction.
            builder.EndTable();

            // Split the table after the 5th row (zero‑based index 5).
            int splitRowIndex = 5;

            // Create a new table that will hold the rows after the split point.
            Table newTable = new Table(doc);

            // Move rows from the original table to the new table.
            // Continue moving while there are rows at the split index.
            while (table.Rows.Count > splitRowIndex)
            {
                // Get the row that should be moved.
                Row rowToMove = table.Rows[splitRowIndex];

                // Detach the row from the original table.
                rowToMove.Remove();

                // Append the detached row to the new table.
                newTable.Rows.Add(rowToMove);
            }

            // Insert the newly created table immediately after the original one.
            table.ParentNode.InsertAfter(newTable, table);

            // Save the resulting document.
            doc.Save("TableSplitResult.docx");
        }
    }
}
