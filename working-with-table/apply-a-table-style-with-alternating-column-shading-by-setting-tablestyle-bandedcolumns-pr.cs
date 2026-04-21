using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleColumnBandingExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // Insert the first cell (required before setting any table formatting).
            builder.InsertCell();

            // Apply a built‑in table style.
            table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

            // Enable column banding (alternating column shading).
            table.StyleOptions = TableStyleOptions.ColumnBands;

            // Build a simple 3 × 3 table.
            for (int row = 0; row < 3; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    builder.Write($"R{row + 1}C{col + 1}");
                    builder.InsertCell();
                }
                builder.EndRow();
                // Start a new row after the first cell has been inserted.
                if (row < 2) builder.InsertCell();
            }

            // Finish the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleColumnBanding.docx");
            doc.Save(outputPath);
        }
    }
}
