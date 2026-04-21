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

            // Build a sample table with 5 rows and 2 columns.
            Table table = builder.StartTable();
            for (int row = 1; row <= 5; row++)
            {
                // First cell of the row.
                builder.InsertCell();
                builder.Write($"Row {row}, Cell 1");

                // Second cell of the row.
                builder.InsertCell();
                builder.Write($"Row {row}, Cell 2");

                // End the current row.
                builder.EndRow();
            }
            // Finish the table.
            builder.EndTable();

            // The table we just created is the first (and only) table in the document.
            Table firstTable = doc.FirstSection.Body.Tables[0];

            // Split the table at the third row (zero‑based index 2).
            int splitRowIndex = 2;

            // Create a new table that will contain the rows from splitRowIndex onward.
            Table secondTable = new Table(doc);
            // Insert the new table right after the original one.
            firstTable.ParentNode.InsertAfter(secondTable, firstTable);

            // Move rows from the original table to the new table.
            while (firstTable.Rows.Count > splitRowIndex)
            {
                // Get the row that should be moved.
                Row movingRow = firstTable.Rows[splitRowIndex];
                // Detach it from the original table.
                movingRow.Remove();
                // Append it to the new table.
                secondTable.Rows.Add(movingRow);
            }

            // Validate that the split produced two tables with the expected row counts.
            if (doc.FirstSection.Body.Tables.Count != 2)
                throw new InvalidOperationException("The document should contain exactly two tables after splitting.");

            if (firstTable.Rows.Count != splitRowIndex)
                throw new InvalidOperationException($"The first table should have {splitRowIndex} rows after splitting.");

            int expectedSecondRows = 5 - splitRowIndex;
            if (secondTable.Rows.Count != expectedSecondRows)
                throw new InvalidOperationException($"The second table should have {expectedSecondRows} rows after splitting.");

            // Save the resulting document.
            doc.Save("SplitTable.docx");
        }
    }
}
