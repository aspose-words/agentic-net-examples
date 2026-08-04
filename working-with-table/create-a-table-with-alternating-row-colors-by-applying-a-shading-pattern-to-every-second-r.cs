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

        // Start the table.
        builder.StartTable();

        int rowCount = 6;   // Number of rows to create.
        int colCount = 3;   // Number of columns per row.

        for (int i = 0; i < rowCount; i++)
        {
            // Populate the cells of the current row.
            for (int j = 0; j < colCount; j++)
            {
                builder.InsertCell();
                builder.Write($"Row {i + 1}, Col {j + 1}");
            }

            // End the current row and obtain the Row object.
            Row row = builder.EndRow();

            // Apply shading to every second row (i.e., rows with odd index).
            if (i % 2 == 1)
            {
                foreach (Cell cell in row.Cells)
                {
                    cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                }
            }
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends automatically; no user interaction required.
    }
}
