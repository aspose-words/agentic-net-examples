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

        int rowCount = 5;
        int columnCount = 4;

        // Start the table.
        builder.StartTable();

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < columnCount; col++)
            {
                // Insert a new cell.
                builder.InsertCell();

                // Apply shading to every second column (1‑based even columns).
                if (col % 2 == 1)
                {
                    // The current cell is the parent of the paragraph the builder is in.
                    Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
                    cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }

                // Write some sample text into the cell.
                builder.Writeln($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "AlternatingColumnShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // The program finishes here.
    }
}
