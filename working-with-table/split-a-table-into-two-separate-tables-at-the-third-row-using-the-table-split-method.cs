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

            // Build a sample table with 6 rows and 2 columns.
            Table table = builder.StartTable();

            for (int i = 1; i <= 6; i++)
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

            // Split the table after the third row (keep first three rows in the original table).
            const int splitRowIndex = 3; // zero‑based index where the second table starts

            // Create a new table that will hold the rows after the split point.
            Table secondTable = new Table(doc);
            // Insert the new table right after the original one in the document tree.
            table.ParentNode.InsertAfter(secondTable, table);

            // Move rows from the original table to the new table.
            // Start from the last row and move upwards to avoid index shifting.
            for (int i = table.Rows.Count - 1; i >= splitRowIndex; i--)
            {
                Row row = table.Rows[i];
                row.Remove();               // Detach from the original table.
                secondTable.Rows.Add(row);   // Append to the new table.
            }

            // Verify the split result.
            Console.WriteLine($"Original table rows after split: {table.Rows.Count}");
            Console.WriteLine($"Second table rows after split: {secondTable.Rows.Count}");

            // Save the document to a file.
            doc.Save("SplitTable.docx");
        }
    }
}
