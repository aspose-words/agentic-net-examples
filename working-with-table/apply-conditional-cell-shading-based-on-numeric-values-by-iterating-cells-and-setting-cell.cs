using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ConditionalCellShadingExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with numeric values.
        // The table will have 5 rows and 3 columns.
        builder.StartTable();

        for (int row = 0; row < 5; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                // Insert a cell and write a numeric value.
                builder.InsertCell();
                int value = (row + 1) * (col + 1) * 10; // Example values: 10,20,...,150
                builder.Write(value.ToString());
            }
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = doc.FirstSection.Body.Tables[0];

        // Iterate through each cell, parse its numeric content, and apply shading based on the value.
        foreach (Row tableRow in table.Rows)
        {
            foreach (Cell cell in tableRow.Cells)
            {
                // Extract the cell text and try to parse it as an integer.
                string cellText = cell.GetText().Trim(); // GetText includes a cell end marker, so trim whitespace.
                if (int.TryParse(cellText, out int numericValue))
                {
                    // Apply green shading for values greater than or equal to 50,
                    // otherwise apply a light salmon shading.
                    if (numericValue >= 50)
                    {
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGreen;
                    }
                    else
                    {
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.LightSalmon;
                    }
                }
            }
        }

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ConditionalCellShading.docx");
        doc.Save(outputPath);
    }
}
