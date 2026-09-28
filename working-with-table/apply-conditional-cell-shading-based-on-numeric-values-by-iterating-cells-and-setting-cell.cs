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

        // Build a simple table with a header row and numeric values.
        builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();

        // Data rows with numeric values.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln((i * 10).ToString()); // Values: 10,20,30,40,50
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Apply conditional shading: cells with numeric value > 25 get yellow background.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Extract cell text and try to parse it as a number.
                string cellText = cell.GetText().Trim();
                if (double.TryParse(cellText, out double numericValue))
                {
                    if (numericValue > 25)
                    {
                        // Set solid yellow shading for the cell by modifying the existing Shading object.
                        cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
                        cell.CellFormat.Shading.ForegroundPatternColor = Color.Yellow;
                    }
                }
            }
        }

        // Save the document to disk.
        string outputPath = "ConditionalShading.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("The output document was not saved correctly.");
        }
    }
}
