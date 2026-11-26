using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableTraversal
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build first sample table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1");
            builder.InsertCell();
            builder.Write("Row 1, Cell 2");
            builder.EndRow();

            builder.InsertCell();
            builder.Write("Row 2, Cell 1");
            builder.InsertCell();
            builder.Write("Row 2, Cell 2");
            builder.EndRow();
            builder.EndTable();

            // Add a paragraph between tables.
            builder.Writeln();

            // Build second sample table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Only Cell");
            builder.EndRow();
            builder.EndTable();

            // Save the sample document.
            const string outputPath = "Sample.docx";
            doc.Save(outputPath);

            // Retrieve all tables in the document.
            NodeCollection tableNodes = doc.GetChildNodes(NodeType.Table, true);

            // Output information about each table.
            int tableIndex = 1;
            foreach (Table table in tableNodes)
            {
                int rowCount = table.Rows.Count;
                int columnCount = rowCount > 0 ? table.FirstRow.Cells.Count : 0;
                Console.WriteLine($"Table {tableIndex}: {rowCount} rows, {columnCount} columns");
                tableIndex++;
            }
        }
    }
}
