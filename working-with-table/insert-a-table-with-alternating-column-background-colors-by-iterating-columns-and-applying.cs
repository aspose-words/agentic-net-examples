using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Drawing;

namespace AsposeWordsTableExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define table dimensions.
            int rows = 4;
            int columns = 5;

            // Start building the table.
            Table table = builder.StartTable();

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    // Insert a new cell.
                    builder.InsertCell();

                    // Apply alternating background colors based on column index.
                    // Even columns get LightGray, odd columns get White.
                    if (col % 2 == 0)
                        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
                    else
                        builder.CellFormat.Shading.BackgroundPatternColor = Color.White;

                    // Write some sample text into the cell.
                    builder.Write($"R{row}C{col}");
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnsTable.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not created.");

            // The program ends automatically; no user interaction required.
        }
    }
}
