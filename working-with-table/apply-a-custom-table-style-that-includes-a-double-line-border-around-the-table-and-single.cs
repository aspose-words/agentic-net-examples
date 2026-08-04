using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a 3x3 table.
        Table table = builder.StartTable();

        // First row
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.InsertCell();
        builder.Write("R1C3");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.InsertCell();
        builder.Write("R2C3");
        builder.EndRow();

        // Third row
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.InsertCell();
        builder.Write("R3C3");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a double line border around the whole table.
        table.ClearBorders(); // Remove any existing borders.
        table.SetBorder(BorderType.Left,   LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

        // Apply single line borders to the inside of the table (each cell).
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Borders.LineStyle = LineStyle.Single;
                cell.CellFormat.Borders.Color = Color.Black;
                cell.CellFormat.Borders.LineWidth = 1.0;
            }
        }

        // Save the document.
        string outputPath = "CustomTableStyle.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output file was not created.");
    }
}
