using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace TableAutoFitExample
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

            // First row, first cell.
            builder.InsertCell();
            // Fix column width to 100 points.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Short");

            // First row, second cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Short");
            builder.EndRow();

            // Second row, first cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Short");

            // Second row, second cell.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Write("Short");
            builder.EndTable();

            // Disable automatic resizing (AutoFit) for the table.
            table.AllowAutoFit = false;
            // Alternatively, you can use:
            // table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // After the table is built, replace the content of the first cell with a long string.
            Cell firstCell = table.FirstRow.FirstCell;
            firstCell.RemoveAllChildren(); // Clear existing content.
            firstCell.AppendChild(new Paragraph(doc));
            firstCell.FirstParagraph.AppendChild(new Run(doc, "This is a very long piece of text that would normally cause the column to expand, but AutoFit is disabled, so the text will be clipped or wrapped according to the cell's settings."));

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAutoFitDisabled.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new Exception("Failed to create the output document.");

            // Informative message (optional, not required for the task).
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
