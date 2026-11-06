using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace ConditionalFormattingExample
{
    public class Program
    {
        public static void Main()
        {
            // Define output file path.
            string outputPath = "ConditionalFormatting.docx";

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table with 5 rows and 3 columns.
            Table table = builder.StartTable();

            // Header row.
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"Header {col + 1}");
                builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;
            }
            builder.EndRow();

            // Data rows with numeric values.
            Random rnd = new Random();
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.InsertCell();
                    // Generate a random integer between 10 and 100.
                    int value = rnd.Next(10, 101);
                    builder.Writeln(value.ToString());
                }
                builder.EndRow();
            }

            builder.EndTable();

            // Define the threshold for highlighting.
            double threshold = 50.0;

            // Apply conditional formatting: highlight cells with values above the threshold.
            foreach (Row row in table.Rows)
            {
                // Skip header row (optional, based on content).
                foreach (Cell cell in row.Cells)
                {
                    // Try to parse the cell text as a double.
                    if (double.TryParse(cell.GetText().Trim('\a', '\r', '\n'), out double numericValue))
                    {
                        if (numericValue > threshold)
                        {
                            // Apply a light yellow background to highlight the cell.
                            cell.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightYellow;
                        }
                    }
                }
            }

            // Save the document to disk.
            doc.Save(outputPath);
        }
    }
}
