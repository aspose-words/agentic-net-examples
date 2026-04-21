using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

namespace AsposeWordsTableFromHtml
{
    public class Program
    {
        public static void Main()
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // HTML string that contains a table with cell formatting (background color and width).
            const string html = @"
                <table style='border-collapse:collapse; border:1px solid black;'>
                    <tr>
                        <td style='background:#FFCCCC; width:120pt; padding:5pt;'>Red Cell</td>
                        <td style='background:#CCFFCC; width:80pt; padding:5pt;'>Green Cell</td>
                    </tr>
                    <tr>
                        <td style='background:#CCCCFF; width:120pt; padding:5pt;'>Blue Cell</td>
                        <td style='background:#FFFFCC; width:80pt; padding:5pt;'>Yellow Cell</td>
                    </tr>
                </table>";

            // Insert the HTML fragment while preserving block‑level formatting.
            builder.InsertHtml(html, HtmlInsertOptions.PreserveBlocks);

            // Save the document to a local file.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromHtml.docx");
            doc.Save(outputPath, SaveFormat.Docx);

            // Verify that the table was inserted and that cell shading matches the HTML.
            Table table = doc.FirstSection.Body.Tables[0];
            if (table == null || table.Rows.Count != 2 || table.Rows[0].Cells.Count != 2)
                throw new InvalidOperationException("The expected table structure was not found.");

            // Helper to compare colors (ignoring alpha channel).
            static bool ColorsEqual(Color a, Color b) => a.R == b.R && a.G == b.G && a.B == b.B;

            // Expected colors from the HTML.
            var expectedColors = new[]
            {
                ColorTranslator.FromHtml("#FFCCCC"),
                ColorTranslator.FromHtml("#CCFFCC"),
                ColorTranslator.FromHtml("#CCCCFF"),
                ColorTranslator.FromHtml("#FFFFCC")
            };

            // Validate each cell's background shading.
            int index = 0;
            foreach (Row row in table.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    Color actual = cell.CellFormat.Shading.BackgroundPatternColor;
                    if (!ColorsEqual(actual, expectedColors[index]))
                        throw new InvalidOperationException($"Cell {index + 1} shading does not match expected color.");
                    index++;
                }
            }

            // If we reach this point, the table was inserted and formatting preserved.
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
    }
}
