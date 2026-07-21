using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableSplitExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with 9 rows and 2 columns.
            Table table = builder.StartTable();

            for (int row = 1; row <= 9; row++)
            {
                // First column.
                builder.InsertCell();
                builder.Write($"Row {row}, Column 1");

                // Second column.
                builder.InsertCell();
                builder.Write($"Row {row}, Column 2");

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // The original table now contains all 9 rows.
            Table firstPart = doc.FirstSection.Body.Tables[0];

            // Create two new empty tables that will hold the split parts.
            Table secondPart = new Table(doc);
            Table thirdPart = new Table(doc);

            // ----- Move rows 4‑6 to the second table -----
            for (int i = 0; i < 3; i++)
            {
                // Rows are zero‑based; row index 3 is the fourth row.
                Row movingRow = firstPart.Rows[3];
                movingRow.Remove();               // Detach from the original table.
                secondPart.Rows.Add(movingRow);   // Add to the second table.
            }

            // Insert the second table right after the first one in the document body.
            firstPart.ParentNode.InsertAfter(secondPart, firstPart);

            // ----- Move rows 7‑9 to the third table -----
            for (int i = 0; i < 3; i++)
            {
                // After the previous removal, the fourth row (index 3) is now the next row to move.
                Row movingRow = firstPart.Rows[3];
                movingRow.Remove();
                thirdPart.Rows.Add(movingRow);
            }

            // Insert the third table after the second one.
            secondPart.ParentNode.InsertAfter(thirdPart, secondPart);

            // Verify that we now have three separate tables.
            int tableCount = doc.FirstSection.Body.Tables.Count;
            Console.WriteLine($"Number of tables after splitting: {tableCount}");

            // Output the row count of each table for confirmation.
            Console.WriteLine($"Table 1 rows: {firstPart.Rows.Count}");
            Console.WriteLine($"Table 2 rows: {secondPart.Rows.Count}");
            Console.WriteLine($"Table 3 rows: {thirdPart.Rows.Count}");

            // Save the resulting document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SplitTable.docx");
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
