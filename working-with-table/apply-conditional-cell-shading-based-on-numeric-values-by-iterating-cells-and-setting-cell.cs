using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class ConditionalCellShading
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x3 table with numeric values.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Score");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apple");
        builder.InsertCell();
        builder.Write("12");
        builder.InsertCell();
        builder.Write("8");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Banana");
        builder.InsertCell();
        builder.Write("25");
        builder.InsertCell();
        builder.Write("15");
        builder.EndRow();

        // Third data row.
        builder.InsertCell();
        builder.Write("Cherry");
        builder.InsertCell();
        builder.Write("7");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Iterate through all cells (excluding header) and apply shading based on numeric values.
        // Cells in the second and third columns contain numbers.
        for (int rowIndex = 1; rowIndex < table.Rows.Count; rowIndex++) // start from 1 to skip header
        {
            Row row = table.Rows[rowIndex];
            for (int cellIndex = 1; cellIndex < row.Cells.Count; cellIndex++) // columns with numbers
            {
                Cell cell = row.Cells[cellIndex];
                // Extract the numeric text from the cell.
                string cellText = cell.ToString(SaveFormat.Text).Trim();

                if (int.TryParse(cellText, out int value))
                {
                    // Apply shading: values > 15 get LightSalmon, otherwise LightGreen.
                    if (value > 15)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightSalmon;
                    else
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                }
            }
        }

        // Define output path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "ConditionalCellShading.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
