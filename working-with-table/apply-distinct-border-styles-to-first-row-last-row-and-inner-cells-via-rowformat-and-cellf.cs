using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableBorderDemo
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

            // ---------- First Row ----------
            // Apply a thick red border to the first row via RowFormat.
            builder.RowFormat.Borders.LineStyle = LineStyle.Single;
            builder.RowFormat.Borders.LineWidth = 2.0; // points
            builder.RowFormat.Borders.Color = Color.Red;

            // Create three cells for the first row.
            builder.InsertCell();
            builder.Write("First Row, Cell 1");
            builder.InsertCell();
            builder.Write("First Row, Cell 2");
            builder.InsertCell();
            builder.Write("First Row, Cell 3");
            builder.EndRow();

            // ---------- Middle Row(s) ----------
            // Clear the previous row formatting so middle rows are not affected.
            builder.RowFormat.ClearFormatting();

            // Apply a thin green border to each inner cell via CellFormat.
            builder.CellFormat.Borders.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.LineWidth = 0.5; // points
            builder.CellFormat.Borders.Color = Color.Green;

            // Create a middle row with the same number of cells.
            builder.InsertCell();
            builder.Write("Middle Row, Cell 1");
            builder.InsertCell();
            builder.Write("Middle Row, Cell 2");
            builder.InsertCell();
            builder.Write("Middle Row, Cell 3");
            builder.EndRow();

            // ---------- Last Row ----------
            // Clear cell formatting to avoid inheriting the green borders.
            builder.CellFormat.ClearFormatting();

            // Apply a thick blue border to the last row via RowFormat.
            builder.RowFormat.Borders.LineStyle = LineStyle.Single;
            builder.RowFormat.Borders.LineWidth = 2.0; // points
            builder.RowFormat.Borders.Color = Color.Blue;

            // Create three cells for the last row.
            builder.InsertCell();
            builder.Write("Last Row, Cell 1");
            builder.InsertCell();
            builder.Write("Last Row, Cell 2");
            builder.InsertCell();
            builder.Write("Last Row, Cell 3");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Save the document to the current directory.
            string outputPath = "TableBordersDemo.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

            // Inform the user (no interactive pause required).
            Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
    }
}
