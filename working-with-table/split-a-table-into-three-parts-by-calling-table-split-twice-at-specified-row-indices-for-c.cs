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

            // Build a table with 9 rows, each row containing a single cell with text.
            Table table = builder.StartTable();
            for (int i = 1; i <= 9; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i}");
                builder.EndRow();
            }
            builder.EndTable();

            // Clone the original table twice – these clones will become the second and third parts.
            Table secondTable = (Table)table.Clone(true);
            Table thirdTable = (Table)table.Clone(true);

            // Insert the cloned tables after the original table so the document contains three tables in order.
            // The ParentNode of a Table is a Body, which derives from CompositeNode and provides InsertAfter<T>.
            var parent = (CompositeNode)table.ParentNode;
            parent.InsertAfter(secondTable, table);
            parent.InsertAfter(thirdTable, secondTable);

            // Keep only the required rows in each table.
            // First part: rows 0‑2
            KeepRows(table, 0, 2);
            // Second part: rows 3‑5
            KeepRows(secondTable, 3, 5);
            // Third part: rows 6‑8
            KeepRows(thirdTable, 6, 8);

            // Optional: Verify the row counts of each part.
            Console.WriteLine($"First part rows: {table.Rows.Count}");
            Console.WriteLine($"Second part rows: {secondTable.Rows.Count}");
            Console.WriteLine($"Third part rows: {thirdTable.Rows.Count}");

            // Save the document containing the three split tables.
            doc.Save("SplitTable.docx");
        }

        /// <summary>
        /// Removes all rows from the table except those whose indices are between startIndex and endIndex (inclusive).
        /// </summary>
        private static void KeepRows(Table tbl, int startIndex, int endIndex)
        {
            // Remove rows after the desired range.
            for (int i = tbl.Rows.Count - 1; i > endIndex; i--)
                tbl.Rows.RemoveAt(i);

            // Remove rows before the desired range.
            for (int i = startIndex - 1; i >= 0; i--)
                tbl.Rows.RemoveAt(i);
        }
    }
}
