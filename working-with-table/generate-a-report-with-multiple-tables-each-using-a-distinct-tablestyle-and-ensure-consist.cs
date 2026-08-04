using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeTablesReport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define a consistent spacing after each table (12 points).
            const double spaceAfterTable = 12.0;

            // Create three tables, each with a distinct built‑in style.
            for (int tableIndex = 0; tableIndex < 3; tableIndex++)
            {
                // Apply spacing after the table.
                builder.ParagraphFormat.SpaceAfter = spaceAfterTable;

                // Start the table.
                Table table = builder.StartTable();

                // ----- Header row -----
                builder.InsertCell();
                builder.Write($"Table {tableIndex + 1} Header 1");
                builder.InsertCell();
                builder.Write($"Table {tableIndex + 1} Header 2");
                builder.EndRow();

                // ----- Data rows -----
                for (int row = 1; row <= 3; row++)
                {
                    builder.InsertCell();
                    builder.Write($"Row {row} Col 1");
                    builder.InsertCell();
                    builder.Write($"Row {row} Col 2");
                    builder.EndRow();
                }

                // Finish the table.
                table = builder.EndTable();

                // Apply a distinct built‑in style to each table.
                switch (tableIndex)
                {
                    case 0:
                        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
                        break;
                    case 1:
                        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;
                        break;
                    case 2:
                        table.StyleIdentifier = StyleIdentifier.TableGrid;
                        break;
                }

                // Apply style options (first row as header, row banding).
                table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

                // Ensure consistent cell spacing inside the table.
                table.AllowCellSpacing = true;
                table.CellSpacing = 5.0;

                // Reset paragraph spacing so it does not affect following content.
                builder.ParagraphFormat.SpaceAfter = 0;
            }

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ReportWithMultipleTables.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("The report file was not created.");

            // Optionally, inform the user (no interactive pause required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
