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

        builder.EndTable();

        // Apply a custom border color (e.g., Blue) to every cell in the first column.
        foreach (Row tableRow in table.Rows)
        {
            Cell firstCell = tableRow.FirstCell;
            // Set all four borders of the cell to the desired color.
            firstCell.CellFormat.Borders[BorderType.Left].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Right].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Top].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Bottom].Color = Color.Blue;
        }

        // Define output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstColumnBorderColor.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
