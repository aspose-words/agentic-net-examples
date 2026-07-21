using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 3x3 table.
            Table table = builder.StartTable();

            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Set a uniform width (in points) for every cell in each column.
            const double uniformWidth = 100.0; // points

            // Determine the number of columns from the first row.
            int columnCount = table.FirstRow.Cells.Count;

            foreach (Row tableRow in table.Rows)
            {
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    Cell cell = tableRow.Cells[colIndex];
                    cell.CellFormat.Width = uniformWidth;
                }
            }

            // Optional: disable auto‑fit to keep the fixed widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "UniformColumnWidths.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
