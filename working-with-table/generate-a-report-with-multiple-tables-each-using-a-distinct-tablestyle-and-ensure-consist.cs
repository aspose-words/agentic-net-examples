using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableReport
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Define three different built‑in table styles to use.
            StyleIdentifier[] styles = new[]
            {
                StyleIdentifier.LightShadingAccent1,
                StyleIdentifier.MediumShading1Accent1,
                StyleIdentifier.ColorfulShadingAccent1
            };

            // Build three tables, each with a distinct style.
            for (int i = 0; i < styles.Length; i++)
            {
                // Ensure a consistent space after the previous table.
                builder.ParagraphFormat.SpaceAfter = 12;
                builder.Writeln(); // blank paragraph that carries the spacing.

                // Start a new table.
                Table table = builder.StartTable();

                // Insert the first cell – this creates the first row, which allows us to set style.
                builder.InsertCell();

                // Apply a distinct style to the current table.
                table.StyleIdentifier = styles[i];
                table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;
                table.AutoFit(AutoFitBehavior.AutoFitToContents);

                // Header row.
                builder.Writeln($"Header {i + 1} - Col 1");
                builder.InsertCell();
                builder.Writeln($"Header {i + 1} - Col 2");
                builder.EndRow();

                // First data row.
                builder.InsertCell();
                builder.Writeln($"Row 1, Cell 1 (Table {i + 1})");
                builder.InsertCell();
                builder.Writeln($"Row 1, Cell 2 (Table {i + 1})");
                builder.EndRow();

                // Second data row.
                builder.InsertCell();
                builder.Writeln($"Row 2, Cell 1 (Table {i + 1})");
                builder.InsertCell();
                builder.Writeln($"Row 2, Cell 2 (Table {i + 1})");
                builder.EndRow();

                // Finish the current table.
                builder.EndTable();

                // Add a blank paragraph after the table to keep spacing consistent.
                builder.Writeln();
            }

            // Save the document.
            string outputPath = "ReportWithMultipleTables.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
