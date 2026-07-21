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

        // Start a table.
        Table table = builder.StartTable();

        // Build a table with 6 rows and 3 columns.
        for (int row = 0; row < 6; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Write($"Row {row + 1}, Cell {col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply alternating background colors to rows.
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
        {
            // Even rows get LightGray, odd rows get White.
            Color background = (rowIndex % 2 == 0) ? Color.LightGray : Color.White;

            foreach (Cell cell in table.Rows[rowIndex].Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = background;
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
