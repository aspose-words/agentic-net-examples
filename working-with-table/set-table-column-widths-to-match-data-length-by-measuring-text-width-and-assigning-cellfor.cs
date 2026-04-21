using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnWidthExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Sample data for the table.
            string[,] data = new string[,]
            {
                { "Product", "Description", "Price" },
                { "Apple", "Fresh red apples", "$1.20" },
                { "Banana", "Ripe yellow bananas", "$0.80" },
                { "Cherry", "Sweet cherries from the valley", "$2.50" }
            };

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

            // Build the table with the sample data.
            Table table = builder.StartTable();

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    builder.InsertCell();
                    builder.Write(data[r, c]);
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Determine the required width for each column based on the longest text in that column.
            // Approximate character width: 7 points per character (depends on font, size, etc.).
            // Add a small padding to avoid clipping.
            double[] maxColumnWidths = new double[cols];
            for (int c = 0; c < cols; c++)
            {
                double maxWidth = 0;
                for (int r = 0; r < rows; r++)
                {
                    // Simple measurement: characters * 7 points.
                    double width = data[r, c].Length * 7.0;
                    if (width > maxWidth)
                        maxWidth = width;
                }
                // Add 5 points padding on each side.
                maxColumnWidths[c] = maxWidth + 10.0;
            }

            // Apply the calculated widths to each cell in the corresponding column.
            for (int r = 0; r < rows; r++)
            {
                Row row = table.Rows[r];
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = row.Cells[c];
                    // Use PreferredWidth to set the column width.
                    cell.CellFormat.PreferredWidth = PreferredWidth.FromPoints(maxColumnWidths[c]);
                }
            }

            // Disable auto‑fit so the widths we set are respected.
            table.AllowAutoFit = false;
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableColumnWidths.docx");
            doc.Save(outputPath);

            // Simple verification that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The output document was not saved correctly.");
        }
    }
}
