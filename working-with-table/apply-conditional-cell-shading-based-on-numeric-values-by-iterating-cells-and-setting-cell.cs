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

        // Build a 4x3 table and fill it with sample numeric values.
        int rows = 4;
        int cols = 3;
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                builder.InsertCell();
                int value = (r + 1) * (c + 1) * 10; // sample calculation
                builder.Write(value.ToString());
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Iterate through each cell and apply shading based on its numeric value.
        Table table = doc.FirstSection.Body.Tables[0];
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Retrieve the cell text and try to parse it as an integer.
                string cellText = cell.ToString(SaveFormat.Text).Trim();
                if (int.TryParse(cellText, out int number))
                {
                    // Apply green shading for values >= 50, otherwise apply coral shading.
                    if (number >= 50)
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                    else
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightCoral;
                }
            }
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ConditionalCellShading.docx");
        doc.Save(outputPath);
    }
}
