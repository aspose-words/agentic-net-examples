using System;
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

        int rows = 3;
        int columns = 4;

        // Populate the table with sample data.
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                builder.InsertCell();
                builder.Writeln($"R{r + 1}C{c + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply a built‑in style.
        table.Style = doc.Styles["Table Grid"];

        // Apply alternating column shading manually.
        for (int r = 0; r < table.Rows.Count; r++)
        {
            Row row = table.Rows[r];
            for (int c = 0; c < row.Cells.Count; c++)
            {
                // Shade every other column (e.g., columns 1,3,…).
                if (c % 2 == 0)
                {
                    row.Cells[c].CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = "TableWithAlternatingColumnShading.docx";
        doc.Save(outputPath);

        // Indicate success.
        if (System.IO.File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully: {outputPath}");
        }
    }
}
