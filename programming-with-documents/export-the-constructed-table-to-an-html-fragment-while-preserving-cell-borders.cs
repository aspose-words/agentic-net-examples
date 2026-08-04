using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

namespace AsposeWordsTableToHtml
{
    public class Program
    {
        public static void Main()
        {
            // Define an output folder and ensure it exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a table.
            Table table = builder.StartTable();

            // Apply a uniform border to the whole table (rows and cells).
            builder.RowFormat.Borders.LineStyle = LineStyle.Single;
            builder.RowFormat.Borders.Color = Color.Black;
            builder.RowFormat.Borders.LineWidth = 1.0;

            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Color = Color.Black;
            builder.CellFormat.Borders.LineWidth = 1.0;

            // First row, first cell.
            builder.InsertCell();
            builder.Write("Cell 1,1");

            // First row, second cell.
            builder.InsertCell();
            builder.Write("Cell 1,2");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.Write("Cell 2,1");

            // Second row, second cell.
            builder.InsertCell();
            builder.Write("Cell 2,2");
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Export the constructed table as an HTML fragment.
            string htmlFragment = table.ToString(SaveFormat.Html);

            // Save the HTML fragment to a file.
            string htmlPath = Path.Combine(outputDir, "TableFragment.html");
            File.WriteAllText(htmlPath, htmlFragment);

            // Optional: indicate completion (no interactive input required).
            Console.WriteLine($"HTML fragment saved to: {htmlPath}");
        }
    }
}
