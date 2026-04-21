using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table.
        Table table = builder.StartTable();

        for (int row = 1; row <= 3; row++)
        {
            for (int col = 1; col <= 3; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row}C{col}");
            }
            builder.EndRow();
        }

        // Finish the table and obtain the Table object.
        table = builder.EndTable();

        // -----------------------------------------------------------------
        // Apply a double line border around the whole table.
        // The last parameter (true) removes any existing explicit cell borders.
        // -----------------------------------------------------------------
        table.SetBorder(BorderType.Left,   LineStyle.Double, 2.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 2.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 2.0, System.Drawing.Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, System.Drawing.Color.Black, true);

        // -----------------------------------------------------------------
        // Apply single line borders to the interior of the table.
        // Iterate over every cell and set its borders to a single line.
        // -----------------------------------------------------------------
        foreach (Row r in table.Rows)
        {
            foreach (Cell c in r.Cells)
            {
                c.CellFormat.Borders[BorderType.Left].LineStyle   = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Right].LineStyle  = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Top].LineStyle    = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;

                c.CellFormat.Borders[BorderType.Left].LineWidth   = 1.0;
                c.CellFormat.Borders[BorderType.Right].LineWidth  = 1.0;
                c.CellFormat.Borders[BorderType.Top].LineWidth    = 1.0;
                c.CellFormat.Borders[BorderType.Bottom].LineWidth = 1.0;

                c.CellFormat.Borders[BorderType.Left].Color   = System.Drawing.Color.Black;
                c.CellFormat.Borders[BorderType.Right].Color  = System.Drawing.Color.Black;
                c.CellFormat.Borders[BorderType.Top].Color    = System.Drawing.Color.Black;
                c.CellFormat.Borders[BorderType.Bottom].Color = System.Drawing.Color.Black;
            }
        }

        // Save the document to the local file system.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomTableStyle.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
