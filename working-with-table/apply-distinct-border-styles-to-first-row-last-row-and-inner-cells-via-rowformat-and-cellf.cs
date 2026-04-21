using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace TableBorderDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with 4 rows and 3 columns.
            Table table = builder.StartTable();

            // Populate the table with sample text.
            for (int row = 1; row <= 4; row++)
            {
                for (int col = 1; col <= 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row}C{col}");
                }
                builder.EndRow();
            }

            // Finish building the table.
            builder.EndTable();

            // ----- Apply distinct border styles -----

            // 1. First row – thick red bottom border.
            Row firstRow = table.FirstRow;
            firstRow.RowFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
            firstRow.RowFormat.Borders[BorderType.Bottom].LineWidth = 2.0;
            firstRow.RowFormat.Borders[BorderType.Bottom].Color = Color.Red;

            // 2. Last row – thick blue top border.
            Row lastRow = table.LastRow;
            lastRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
            lastRow.RowFormat.Borders[BorderType.Top].LineWidth = 2.0;
            lastRow.RowFormat.Borders[BorderType.Top].Color = Color.Blue;

            // 3. Inner cells (rows 2 and 3) – thin green left/right borders.
            for (int i = 1; i < table.Rows.Count - 1; i++) // skip first and last rows
            {
                Row innerRow = table.Rows[i];
                foreach (Cell cell in innerRow.Cells)
                {
                    // Use a supported line style (DotDash) for the inner borders.
                    cell.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.DotDash;
                    cell.CellFormat.Borders[BorderType.Left].LineWidth = 1.0;
                    cell.CellFormat.Borders[BorderType.Left].Color = Color.Green;

                    cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.DotDash;
                    cell.CellFormat.Borders[BorderType.Right].LineWidth = 1.0;
                    cell.CellFormat.Borders[BorderType.Right].Color = Color.Green;
                }
            }

            // Save the document to the local file system.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableBorders.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
