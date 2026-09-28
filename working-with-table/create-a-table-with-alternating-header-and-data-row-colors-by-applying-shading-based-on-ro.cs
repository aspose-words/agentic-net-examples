using System;
using System.IO;
using System.Drawing;
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

        // Start the table.
        builder.StartTable();

        // ---------- Header row ----------
        // Apply shading to the header cells.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Insert header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.InsertCell();
        builder.Write("Header 3");
        builder.EndRow();

        // Reset cell shading to default after the header row.
        builder.CellFormat.Shading.Texture = TextureIndex.TextureNone;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;

        // ---------- Data rows ----------
        int dataRowCount = 6; // Number of data rows to create.

        for (int i = 0; i < dataRowCount; i++)
        {
            // Apply alternating shading based on row index parity.
            if (i % 2 == 0) // Even index rows get a light blue background.
            {
                builder.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
            }
            else // Odd index rows have no shading (default background).
            {
                builder.CellFormat.Shading.Texture = TextureIndex.TextureNone;
                builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
            }

            // Insert cells for the current data row.
            builder.InsertCell();
            builder.Write($"Row {i + 1} - Col 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1} - Col 2");
            builder.InsertCell();
            builder.Write($"Row {i + 1} - Col 3");
            builder.EndRow();

            // Reset shading for the next row (optional, will be set again in the loop).
            builder.CellFormat.Shading.Texture = TextureIndex.TextureNone;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.White;
        }

        // End the table.
        builder.EndTable();

        // Save the document to a file.
        string outputPath = "AlternatingRows.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
