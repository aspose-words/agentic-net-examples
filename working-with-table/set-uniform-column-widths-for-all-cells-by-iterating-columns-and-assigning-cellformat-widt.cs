using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3x2 table.
            builder.StartTable();

            // First row.
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R1C{col}");
            }
            builder.EndRow();

            // Second row.
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R2C{col}");
            }
            builder.EndRow();

            builder.EndTable();

            // Retrieve the created table.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
            if (table == null)
                throw new InvalidOperationException("Table was not created.");

            // Determine the number of columns (assumes all rows have the same number of cells).
            int columnCount = table.Rows[0].Cells.Count;

            // Set a uniform width for each column by assigning CellFormat.Width to each cell in the column.
            const double uniformWidth = 100.0; // Width in points.

            for (int colIndex = 0; colIndex < columnCount; colIndex++)
            {
                foreach (Row row in table.Rows)
                {
                    Cell cell = row.Cells[colIndex];
                    cell.CellFormat.Width = uniformWidth;
                }
            }

            // Save the document.
            string outputPath = "UniformColumnWidths.docx";
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new FileNotFoundException("The output document was not created.", outputPath);
        }
    }
}
