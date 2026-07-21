using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableStyleToDirectFormattingDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 2x2 table.
            Table table = builder.StartTable();

            // First row.
            builder.InsertCell();
            builder.Write("Header 1");
            builder.InsertCell();
            builder.Write("Header 2");
            builder.EndRow();

            // Second row.
            builder.InsertCell();
            builder.Write("Cell 1");
            builder.InsertCell();
            builder.Write("Cell 2");
            builder.EndTable();

            // Create a custom table style.
            TableStyle tableStyle = (TableStyle)doc.Styles.Add(StyleType.Table, "MyCustomTableStyle");
            tableStyle.CellSpacing = 5; // Space between cells.
            tableStyle.Shading.BackgroundPatternColor = Color.AntiqueWhite; // Table background.
            tableStyle.Borders.Color = Color.Blue; // Border color.
            tableStyle.Borders.LineStyle = LineStyle.DotDash; // Border style.
            tableStyle.RowStripe = 2; // Example of row banding (odd/even rows).

            // Apply the style to the table.
            table.Style = tableStyle;

            // Convert the style formatting into direct formatting on the table elements.
            doc.ExpandTableStylesToDirectFormatting();

            // ----- Validation (optional) -----
            // After expansion, the style properties should be reflected directly on the table and its cells.
            Console.WriteLine("Table CellSpacing (direct): " + table.CellSpacing);
            Console.WriteLine("First cell background color (direct): " +
                table.FirstRow.FirstCell.CellFormat.Shading.BackgroundPatternColor.Name);
            Console.WriteLine("First cell left border color (direct): " +
                table.FirstRow.FirstCell.CellFormat.Borders[BorderType.Left].Color.Name);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the resulting document.
            string outputPath = Path.Combine(outputDir, "TableStyleExpanded.docx");
            doc.Save(outputPath);

            // Indicate completion.
            Console.WriteLine("Document saved to: " + outputPath);
        }
    }
}
