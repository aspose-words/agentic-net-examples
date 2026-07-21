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

        // Build a simple 3x2 table.
        Table table = builder.StartTable();

        // Row 1
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Write("R3C1");
        builder.InsertCell();
        builder.Write("R3C2");
        builder.EndRow();

        // Finish the table and obtain the Table object.
        table = builder.EndTable();

        // Apply a custom border color (e.g., Blue) to all cells in the first column.
        foreach (Row row in table.Rows)
        {
            Cell firstCell = row.FirstCell;
            // Set the left border color.
            firstCell.CellFormat.Borders[BorderType.Left].Color = Color.Blue;
            // Optionally set other borders of the first column cells to the same color.
            firstCell.CellFormat.Borders[BorderType.Top].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Bottom].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Right].Color = Color.Blue;
        }

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstColumnBorder.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
