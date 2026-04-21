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
        Table table = builder.StartTable();

        // Define number of rows and columns.
        int rows = 6;      // 1 header row + 5 data rows
        int columns = 3;

        for (int rowIndex = 0; rowIndex < rows; rowIndex++)
        {
            // Choose shading color based on row index parity.
            Color shadingColor;
            if (rowIndex == 0)                     // Header row
                shadingColor = Color.LightBlue;
            else if (rowIndex % 2 == 0)            // Even data rows
                shadingColor = Color.LightGray;
            else                                   // Odd data rows
                shadingColor = Color.White;

            // Apply shading to the cells that will be created in this row.
            builder.CellFormat.Shading.BackgroundPatternColor = shadingColor;

            // Populate cells for the current row.
            for (int colIndex = 0; colIndex < columns; colIndex++)
            {
                builder.InsertCell();
                builder.Write($"Row {rowIndex}, Col {colIndex}");
            }

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRowsTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // The program ends automatically; no user interaction required.
    }
}
