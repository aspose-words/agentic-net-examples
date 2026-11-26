using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Add header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Add data rows.
        for (int i = 1; i <= 6; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.InsertCell();
            builder.Write($"Row {i} Col 3");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply alternating shading (light gray) to every other data row (starting with the first data row).
        // Row index 0 is the header row, so we start shading from index 1.
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++)
        {
            // Shade odd-numbered rows (1,3,5, ...) relative to the data rows.
            if ((rowIndex - 1) % 2 == 0)
            {
                Row row = table.Rows[rowIndex];
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                }
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRowsTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created and can be reloaded.
        if (File.Exists(outputPath))
        {
            Document loadedDoc = new Document(outputPath);
            // Document successfully reloaded; further processing could be added here.
        }
    }
}
