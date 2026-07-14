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

            // Build a sample table with 5 rows and 2 columns.
            Table table = builder.StartTable();
            for (int i = 1; i <= 5; i++)
            {
                // First cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");

                // Second cell of the row.
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");

                // End the current row.
                builder.EndRow();
            }
            // Finish the table construction.
            builder.EndTable();

            // Split the table so that the first three rows stay in the original table
            // and the remaining rows move to a new table.
            int splitIndex = 3; // zero‑based index; rows with index >= 3 will be moved.
            Table newTable = new Table(doc);

            // Move rows from the original table to the new table.
            while (table.Rows.Count > splitIndex)
            {
                // The row at splitIndex is the first row that should be moved.
                Row rowToMove = table.Rows[splitIndex];
                rowToMove.Remove();               // Detach from the original table.
                newTable.Rows.Add(rowToMove);      // Append to the new table.
            }

            // Insert the new table immediately after the original table in the document.
            table.ParentNode.InsertAfter(newTable, table);

            // Optional validation (can be removed if not needed).
            Console.WriteLine($"Original table rows after split: {table.Rows.Count}"); // Expected: 3
            Console.WriteLine($"New table rows after split: {newTable.Rows.Count}");   // Expected: 2

            // Save the document containing the two separate tables.
            doc.Save("SplitTable.docx");
        }
    }
}
