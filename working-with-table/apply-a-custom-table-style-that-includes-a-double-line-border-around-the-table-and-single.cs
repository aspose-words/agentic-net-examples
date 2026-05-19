using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 3x3 table with sample text.
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

        // Apply double line borders to the outer edges of the table.
        table.SetBorder(BorderType.Left,   LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);

        // Apply single line borders to the interior cell edges.
        foreach (Row r in table.Rows)
        {
            foreach (Cell c in r.Cells)
            {
                // Left border
                c.CellFormat.Borders[BorderType.Left].LineStyle = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Left].LineWidth = 1.0;
                c.CellFormat.Borders[BorderType.Left].Color = Color.Black;

                // Right border
                c.CellFormat.Borders[BorderType.Right].LineStyle = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Right].LineWidth = 1.0;
                c.CellFormat.Borders[BorderType.Right].Color = Color.Black;

                // Top border
                c.CellFormat.Borders[BorderType.Top].LineStyle = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Top].LineWidth = 1.0;
                c.CellFormat.Borders[BorderType.Top].Color = Color.Black;

                // Bottom border
                c.CellFormat.Borders[BorderType.Bottom].LineStyle = LineStyle.Single;
                c.CellFormat.Borders[BorderType.Bottom].LineWidth = 1.0;
                c.CellFormat.Borders[BorderType.Bottom].Color = Color.Black;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTableStyle.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
