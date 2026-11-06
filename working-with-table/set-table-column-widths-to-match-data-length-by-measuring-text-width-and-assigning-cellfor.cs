using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableColumnWidthExample
{
    public class Program
    {
        // Simple text width estimator: character count * font size * factor.
        private static double EstimateTextWidth(string text, double fontSize)
        {
            const double factor = 1.2; // Approximate width per point size.
            return text.Length * fontSize * factor;
        }

        public static void Main()
        {
            // Sample data for the table.
            string[,] data = new string[,]
            {
                { "ID", "Name", "Occupation" },
                { "1", "Alice", "Software Engineer" },
                { "2", "Bob", "Graphic Designer" },
                { "3", "Charlie", "Data Scientist" }
            };

            // Create a new document and a builder.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build the table using the builder.
            builder.StartTable();

            // Use a consistent font size for measurement.
            builder.Font.Size = 12; // points

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);

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

            // Retrieve the created table.
            Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

            // Determine the maximum required width for each column.
            double[] maxColumnWidths = new double[cols];

            for (int c = 0; c < cols; c++)
            {
                double maxWidth = 0;
                for (int r = 0; r < rows; r++)
                {
                    string cellText = data[r, c];
                    double width = EstimateTextWidth(cellText, builder.Font.Size);
                    if (width > maxWidth)
                        maxWidth = width;
                }
                // Add a small padding to avoid clipping.
                maxColumnWidths[c] = maxWidth + 5; // points
            }

            // Apply the calculated widths to each cell in the corresponding column.
            foreach (Row row in table.Rows)
            {
                for (int c = 0; c < cols; c++)
                {
                    Cell cell = row.Cells[c];
                    cell.CellFormat.Width = maxColumnWidths[c];
                }
            }

            // Save the document.
            string outputPath = "TableColumnWidths.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // The program ends here without waiting for user input.
        }
    }
}
