using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableSplitExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a table with 9 rows and 2 columns.
            Table table = builder.StartTable();
            for (int i = 1; i <= 9; i++)
            {
                builder.InsertCell();
                builder.Write($"Row {i}, Cell 1");

                builder.InsertCell();
                builder.Write($"Row {i}, Cell 2");

                builder.EndRow();
            }
            builder.EndTable();

            // ---- Split the table into three parts ----
            // First split at row index 3 (rows 0‑2 stay in the original table,
            // rows 3‑8 move to the new table).
            Table secondPart = SplitTable(table, 3, table.Rows.Count - 3);

            // Second split the remaining table at row index 3 (relative to this table).
            // This yields the middle part (rows 0‑2 of secondPart) and the last part.
            Table thirdPart = SplitTable(secondPart, 3, secondPart.Rows.Count - 3);

            // Save each part into its own document.
            SaveTablePart(table, "TablePart1.docx");
            SaveTablePart(secondPart, "TablePart2.docx");
            SaveTablePart(thirdPart, "TablePart3.docx");
        }

        /// <summary>
        /// Splits a table by moving a range of rows starting at <paramref name="startIndex"/>
        /// (zero‑based) into a new table. The original table loses those rows.
        /// </summary>
        private static Table SplitTable(Table source, int startIndex, int rowCount)
        {
            // Create a new empty table in the same document as the source.
            Table newTable = new Table(source.Document);

            // Clone the required rows into the new table.
            for (int i = 0; i < rowCount; i++)
            {
                Row clonedRow = (Row)source.Rows[startIndex].Clone(true);
                newTable.Rows.Add(clonedRow);
            }

            // Remove the moved rows from the source table (remove from the end to keep indices stable).
            for (int i = 0; i < rowCount; i++)
            {
                source.Rows.RemoveAt(startIndex);
            }

            return newTable;
        }

        /// <summary>
        /// Creates a new document, imports the supplied table, and saves it.
        /// </summary>
        private static void SaveTablePart(Table sourceTable, string fileName)
        {
            // Create a new document to hold the table.
            Document partDoc = new Document();

            // Import the table into the new document (required when moving nodes between documents).
            Table importedTable = (Table)partDoc.ImportNode(sourceTable, true);
            partDoc.FirstSection.Body.AppendChild(importedTable);

            // Save the document.
            partDoc.Save(fileName);
        }
    }
}
