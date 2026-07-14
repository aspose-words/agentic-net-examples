using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace TableBorderExample
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

            int rows = 4;
            int cols = 3;

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    builder.InsertCell();
                    builder.Write($"R{r}C{c}");
                }
                builder.EndRow();
            }

            // Finish building the table.
            table = builder.EndTable();

            // ---------- Apply distinct border styles ----------

            // 1. First row – double line top and bottom borders, red color.
            Row firstRow = table.FirstRow;
            firstRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.Double;
            firstRow.RowFormat.Borders[BorderType.Top].Color = Color.Red;
            firstRow.RowFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Double;
            firstRow.RowFormat.Borders[BorderType.Bottom].Color = Color.Red;

            // 2. Last row – dash‑dot line top and bottom borders, blue color.
            // The DashDot style is not available in this version of Aspose.Words,
            // so we use DashSmallGap as a visually similar alternative.
            Row lastRow = table.LastRow;
            lastRow.RowFormat.Borders[BorderType.Top].LineStyle = LineStyle.DashSmallGap;
            lastRow.RowFormat.Borders[BorderType.Top].Color = Color.Blue;
            lastRow.RowFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.DashSmallGap;
            lastRow.RowFormat.Borders[BorderType.Bottom].Color = Color.Blue;

            // 3. Inner cells – solid single line borders, green color.
            // Iterate over rows excluding first and last.
            for (int i = 1; i < table.Rows.Count - 1; i++)
            {
                Row innerRow = table.Rows[i];
                foreach (Cell cell in innerRow.Cells)
                {
                    cell.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
                    cell.CellFormat.Borders[BorderType.Left].Color = Color.Green;

                    cell.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
                    cell.CellFormat.Borders[BorderType.Right].Color = Color.Green;

                    cell.CellFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
                    cell.CellFormat.Borders[BorderType.Top].Color = Color.Green;

                    cell.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
                    cell.CellFormat.Borders[BorderType.Bottom].Color = Color.Green;
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "TableBorders.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The document was not saved successfully.");
        }
    }
}
