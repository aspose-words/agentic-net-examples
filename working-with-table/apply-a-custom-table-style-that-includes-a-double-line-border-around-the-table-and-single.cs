using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a 3x3 table.
            Table table = builder.StartTable();

            // Fill the table with sample text.
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    builder.Write($"R{row + 1}C{col + 1}");
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // -----------------------------------------------------------------
            // Apply single line borders to all cell edges (inner grid lines).
            // -----------------------------------------------------------------
            foreach (Row r in table.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    // Set each side of the cell to a single black line.
                    c.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
                    c.CellFormat.Borders[BorderType.Left].LineWidth = 1.0;
                    c.CellFormat.Borders[BorderType.Left].Color = Color.Black;

                    c.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
                    c.CellFormat.Borders[BorderType.Right].LineWidth = 1.0;
                    c.CellFormat.Borders[BorderType.Right].Color = Color.Black;

                    c.CellFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
                    c.CellFormat.Borders[BorderType.Top].LineWidth = 1.0;
                    c.CellFormat.Borders[BorderType.Top].Color = Color.Black;

                    c.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
                    c.CellFormat.Borders[BorderType.Bottom].LineWidth = 1.0;
                    c.CellFormat.Borders[BorderType.Bottom].Color = Color.Black;
                }
            }

            // -----------------------------------------------------------------
            // Apply a double line border around the whole table.
            // The 'isOverrideCellBorders' flag is set to false so that the
            // previously defined cell borders (the inner grid) are preserved.
            // -----------------------------------------------------------------
            table.SetBorder(BorderType.Left, LineStyle.Double, 2.0, Color.Black, false);
            table.SetBorder(BorderType.Right, LineStyle.Double, 2.0, Color.Black, false);
            table.SetBorder(BorderType.Top, LineStyle.Double, 2.0, Color.Black, false);
            table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, false);

            // Save the document to a file.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomTableStyle.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // The program ends here; no user interaction is required.
        }
    }
}
