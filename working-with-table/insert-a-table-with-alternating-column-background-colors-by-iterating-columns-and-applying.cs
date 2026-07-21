using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define the size of the table.
        int rowCount = 4;
        int columnCount = 5;

        // Colors to alternate between columns.
        Color evenColumnColor = Color.LightGray;
        Color oddColumnColor = Color.White;

        // Start building the table.
        Table table = builder.StartTable();

        for (int row = 0; row < rowCount; row++)
        {
            for (int col = 0; col < columnCount; col++)
            {
                // Insert a new cell into the current row.
                builder.InsertCell();

                // Apply shading based on the column index (alternating colors).
                builder.CellFormat.Shading.BackgroundPatternColor = (col % 2 == 0) ? evenColumnColor : oddColumnColor;

                // Add some sample text to the cell.
                builder.Write($"R{row + 1}C{col + 1}");
            }

            // End the current row before starting the next one.
            builder.EndRow();
        }

        // Finish the table construction.
        builder.EndTable();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnsTable.docx");
        doc.Save(outputPath);
    }
}
