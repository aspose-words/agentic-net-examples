using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        int rows = 5;
        int columns = 4;

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                // Apply shading to every second column (1‑based index: columns 2,4,...).
                if (c % 2 == 1)
                {
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
                else
                {
                    // Remove any previous shading.
                    builder.CellFormat.Shading.ClearFormatting();
                }

                // Insert the cell and write some text.
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnShading.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
