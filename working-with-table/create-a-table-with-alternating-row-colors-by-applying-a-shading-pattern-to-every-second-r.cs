using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

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
            int rowCount = 6;      // Number of rows.
            int columnCount = 3;   // Number of columns.

            // Start building the table.
            Table table = builder.StartTable();

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                // Choose background color based on row index (alternating colors).
                Color rowColor = (rowIndex % 2 == 0) ? Color.White : Color.LightGray;

                // Apply shading to the cells that will be created in this row.
                builder.CellFormat.Shading.BackgroundPatternColor = rowColor;

                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {rowIndex + 1}, Col {colIndex + 1}");
                }

                // Finish the current row.
                builder.EndRow();
            }

            // End the table construction.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
