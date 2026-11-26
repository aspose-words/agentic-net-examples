using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Define number of rows and columns.
            int rows = 5;
            int columns = 4;

            // Populate the table with sample data.
            for (int row = 0; row < rows; row++)
            {
                for (int col = 0; col < columns; col++)
                {
                    // Insert cell text.
                    builder.InsertCell();
                    builder.Writeln($"R{row + 1}C{col + 1}");
                }
                // End the current row.
                builder.EndRow();
            }

            // End the table.
            builder.EndTable();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.TableGrid;

            // Enable alternating column shading.
            table.StyleOptions = TableStyleOptions.ColumnBands;

            // Save the document.
            string outputPath = "AlternatingColumnShading.docx";
            doc.Save(outputPath);
        }
    }
}
