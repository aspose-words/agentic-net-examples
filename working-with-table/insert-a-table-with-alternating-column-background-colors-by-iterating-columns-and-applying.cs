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

        // Start building the table.
        Table table = builder.StartTable();

        int rows = 4;
        int columns = 3;

        // Iterate through rows.
        for (int row = 0; row < rows; row++)
        {
            // Iterate through columns.
            for (int col = 0; col < columns; col++)
            {
                // Apply alternating background colors based on column index.
                if (col % 2 == 0)
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                else
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

                // Insert a new cell and write its coordinates.
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnShading.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
