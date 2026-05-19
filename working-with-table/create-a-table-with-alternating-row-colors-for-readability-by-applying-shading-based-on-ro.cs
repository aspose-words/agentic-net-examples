using System;
using System.IO;
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

            // Start the table.
            builder.StartTable();

            // Define number of rows and columns.
            int rowCount = 6;
            int columnCount = 3;

            // Loop through rows and apply alternating shading.
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                // Choose background color based on row parity.
                // Even rows: LightGray, Odd rows: White.
                System.Drawing.Color backgroundColor = (rowIndex % 2 == 0)
                    ? System.Drawing.Color.LightGray
                    : System.Drawing.Color.White;

                // Apply shading to the cells that will be created in this row.
                builder.CellFormat.Shading.BackgroundPatternColor = backgroundColor;

                // Insert cells for the current row.
                for (int colIndex = 0; colIndex < columnCount; colIndex++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {rowIndex + 1}, Col {colIndex + 1}");
                }

                // End the current row.
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Reset cell formatting to avoid affecting subsequent content.
            builder.CellFormat.Shading.ClearFormatting();

            // Save the document to the local file system.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AlternatingRows.docx");
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");
        }
    }
}
