using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableBordersExample
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

            // Insert the first cell where we will apply custom borders.
            builder.InsertCell();

            // Clear any previous cell formatting to start from defaults.
            builder.CellFormat.ClearFormatting();

            // Set different line widths, styles, and colors for each side of the cell border.
            builder.CellFormat.Borders.Left.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Left.LineWidth = 2.0; // points
            builder.CellFormat.Borders.Left.Color = Color.Red;

            builder.CellFormat.Borders.Right.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Right.LineWidth = 4.0;
            builder.CellFormat.Borders.Right.Color = Color.Green;

            builder.CellFormat.Borders.Top.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Top.LineWidth = 6.0;
            builder.CellFormat.Borders.Top.Color = Color.Blue;

            builder.CellFormat.Borders.Bottom.LineStyle = LineStyle.Single;
            builder.CellFormat.Borders.Bottom.LineWidth = 8.0;
            builder.CellFormat.Borders.Bottom.Color = Color.Purple;

            // Add some text to the customized cell.
            builder.Writeln("Cell with custom borders");

            // Insert a second cell with default formatting for comparison.
            builder.InsertCell();
            builder.Writeln("Normal cell");

            // End the first row.
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Define the output path.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "CustomCellBorders.docx");

            // Save the document.
            doc.Save(outputPath);

            // Simple validation to ensure the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // The program finishes here without waiting for user input.
        }
    }
}
