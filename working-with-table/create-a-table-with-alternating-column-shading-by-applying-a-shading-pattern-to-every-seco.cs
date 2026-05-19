using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableShading
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            builder.StartTable();

            int rowCount = 5;
            int columnCount = 4;

            // Build the table with sample text.
            for (int row = 0; row < rowCount; row++)
            {
                for (int col = 0; col < columnCount; col++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {row + 1}, Col {col + 1}");
                }
                builder.EndRow();
            }

            // Finish the table.
            Table table = builder.EndTable();

            // Apply shading to every second column (index 1, 3, ...).
            Color shadingColor = Color.LightGray;
            foreach (Row tableRow in table.Rows)
            {
                for (int colIndex = 0; colIndex < tableRow.Cells.Count; colIndex++)
                {
                    if (colIndex % 2 == 1) // second, fourth, etc.
                    {
                        tableRow.Cells[colIndex].CellFormat.Shading.BackgroundPatternColor = shadingColor;
                    }
                }
            }

            // Define output path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingColumnShading.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved successfully.");
        }
    }
}
