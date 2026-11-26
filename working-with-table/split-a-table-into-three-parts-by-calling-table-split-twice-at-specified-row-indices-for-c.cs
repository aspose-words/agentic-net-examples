using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableSplitExample
{
    class Program
    {
        static void Main()
        {
            // Create a new document and a DocumentBuilder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a sample table with 9 rows and 2 columns.
            builder.StartTable();
            for (int row = 0; row < 9; row++)
            {
                // First cell.
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Cell 1");

                // Second cell.
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Cell 2");

                // End the current row.
                builder.EndRow();
            }
            builder.EndTable();

            // Retrieve the first (and only) table in the document.
            Table firstTable = (Table)doc.GetChild(NodeType.Table, 0, true);

            // ---- First split ----
            // Split the table so that rows 4‑9 (indices 3‑8) become a new table.
            Table secondTable = SplitTable(firstTable, 3, firstTable.Rows.Count - 1);

            // ---- Second split ----
            // Split the second table so that its rows 7‑9 (indices 3‑5) become a new table.
            Table thirdTable = SplitTable(secondTable, 3, secondTable.Rows.Count - 1);

            // Save the resulting document.
            string outputPath = "SplitTable.docx";
            doc.Save(outputPath);

            // Simple validation that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output file was not created.");
        }

        /// <summary>
        /// Splits a table by moving rows from startRowIndex to endRowIndex (inclusive)
        /// into a newly created table that is inserted immediately after the source table.
        /// </summary>
        /// <param name="source">The table to split.</param>
        /// <param name="startRowIndex">Zero‑based index of the first row to move.</param>
        /// <param name="endRowIndex">Zero‑based index of the last row to move.</param>
        /// <returns>The newly created table containing the moved rows.</returns>
        private static Table SplitTable(Table source, int startRowIndex, int endRowIndex)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (startRowIndex < 0 || endRowIndex >= source.Rows.Count || startRowIndex > endRowIndex)
                throw new ArgumentOutOfRangeException("Invalid row indices for split.");

            // Clone the source table without its rows to preserve formatting.
            Table newTable = (Table)source.Clone(false);

            // Move rows from the source to the new table.
            // Iterate backwards to avoid index shifting when removing rows.
            for (int i = endRowIndex; i >= startRowIndex; i--)
            {
                Row row = source.Rows[i];
                Row clonedRow = (Row)row.Clone(true);
                // Insert at the beginning of the new table to keep original order.
                newTable.Rows.Insert(0, clonedRow);
                // Remove the row from the original table.
                source.Rows.RemoveAt(i);
            }

            // Insert the new table right after the source table in the document.
            source.ParentNode.InsertAfter(newTable, source);
            return newTable;
        }
    }
}
