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

        // Define table dimensions.
        int rows = 5;
        int columns = 5;
        int threshold = 50;

        // Insert a table.
        Table table = builder.StartTable();

        // Populate the table with numeric values.
        Random rnd = new Random();
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < columns; j++)
            {
                // Insert a cell with a random integer between 1 and 100.
                builder.InsertCell();
                int value = rnd.Next(1, 101);
                builder.Writeln(value.ToString());
            }
            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Apply conditional formatting: highlight cells with values above the threshold.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Try to parse the cell text to an integer.
                if (int.TryParse(cell.GetText().Trim('\r', '\a'), out int cellValue))
                {
                    if (cellValue > threshold)
                    {
                        // Set background shading to yellow.
                        cell.CellFormat.Shading.BackgroundPatternColor = Color.Yellow;
                    }
                }
            }
        }

        // Save the document.
        string outputPath = "ConditionalFormatting.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (System.IO.File.Exists(outputPath))
        {
            // Optionally, reopen the document to ensure it loads without error.
            Document loadedDoc = new Document(outputPath);
        }
    }
}
