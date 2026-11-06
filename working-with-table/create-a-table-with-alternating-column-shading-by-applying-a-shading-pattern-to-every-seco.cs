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

        // Build a table with 5 rows and 3 columns.
        builder.StartTable();
        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Apply shading to every second column (index 1, 3, ...).
        Table table = doc.FirstSection.Body.Tables[0];
        int columnCount = table.Rows[0].Cells.Count;

        for (int colIndex = 0; colIndex < columnCount; colIndex++)
        {
            if (colIndex % 2 == 1) // every second column
            {
                foreach (Row row in table.Rows)
                {
                    Cell cell = row.Cells[colIndex];
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = "AlternatingColumnShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Program ends automatically.
    }
}
