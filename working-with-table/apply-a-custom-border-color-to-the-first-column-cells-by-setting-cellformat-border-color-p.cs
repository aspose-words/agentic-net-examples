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

        // Start a table and populate it with 3 rows and 3 columns.
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

        // Finish the table.
        builder.EndTable();

        // Apply a custom border color (Blue) to every cell in the first column.
        foreach (Row r in table.Rows)
        {
            Cell firstCell = r.FirstCell;
            firstCell.CellFormat.Borders[BorderType.Left].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Top].Color = Color.Blue;
            firstCell.CellFormat.Borders[BorderType.Bottom].Color = Color.Blue;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstColumnBorderColor.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
