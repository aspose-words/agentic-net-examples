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

        // Define table dimensions.
        int rowCount = 6;
        int columnCount = 3;

        // Build the table.
        Table table = builder.StartTable();
        for (int i = 0; i < rowCount; i++)
        {
            for (int j = 0; j < columnCount; j++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Col {j + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Apply shading to every second row (index 1,3,5,...).
        for (int i = 0; i < table.Rows.Count; i++)
        {
            if (i % 2 == 1) // even-numbered rows in 1‑based terms
            {
                Row row = table.Rows[i];
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
