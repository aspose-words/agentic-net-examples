using System;
using System.IO;
using System.Drawing; // Needed for Color
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTableBordersExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table with 4 rows and 3 columns.
            builder.StartTable();

            // Row 1
            builder.InsertCell();
            builder.Write("R1C1");
            builder.InsertCell();
            builder.Write("R1C2");
            builder.InsertCell();
            builder.Write("R1C3");
            builder.EndRow();

            // Row 2
            builder.InsertCell();
            builder.Write("R2C1");
            builder.InsertCell();
            builder.Write("R2C2");
            builder.InsertCell();
            builder.Write("R2C3");
            builder.EndRow();

            // Row 3
            builder.InsertCell();
            builder.Write("R3C1");
            builder.InsertCell();
            builder.Write("R3C2");
            builder.InsertCell();
            builder.Write("R3C3");
            builder.EndRow();

            // Row 4
            builder.InsertCell();
            builder.Write("R4C1");
            builder.InsertCell();
            builder.Write("R4C2");
            builder.InsertCell();
            builder.Write("R4C3");
            builder.EndRow();

            // Finish the table and obtain a reference to it.
            Table table = builder.EndTable();

            // Apply distinct border styles to the first row using RowFormat.
            Row firstRow = table.FirstRow;
            firstRow.RowFormat.Borders.LineStyle = LineStyle.Double;
            firstRow.RowFormat.Borders.Color = Color.Blue;
            firstRow.RowFormat.Borders.LineWidth = 2.0;

            // Apply distinct border styles to the last row using RowFormat.
            Row lastRow = table.LastRow;
            lastRow.RowFormat.Borders.LineStyle = LineStyle.DashSmallGap; // Replaced unavailable DashDot
            lastRow.RowFormat.Borders.Color = Color.Red;
            lastRow.RowFormat.Borders.LineWidth = 2.0;

            // Apply a different border style to all inner cells using CellFormat.
            for (int i = 1; i < table.Rows.Count - 1; i++) // Skip first and last rows.
            {
                Row innerRow = table.Rows[i];
                foreach (Cell cell in innerRow.Cells)
                {
                    cell.CellFormat.Borders.LineStyle = LineStyle.Single;
                    cell.CellFormat.Borders.Color = Color.Green;
                    cell.CellFormat.Borders.LineWidth = 1.0;
                }
            }

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableBorders.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");
        }
    }
}
