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

        int rows = 4;
        int cols = 4;

        // Build the table rows and cells.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                // Apply shading to every second column (1‑based even columns).
                if (c % 2 == 1)
                    builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                else
                    builder.CellFormat.Shading.ClearFormatting(); // No shading for odd columns.

                // Insert a cell and write some text.
                builder.InsertCell();
                builder.Write($"R{r + 1}C{c + 1}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnsTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
