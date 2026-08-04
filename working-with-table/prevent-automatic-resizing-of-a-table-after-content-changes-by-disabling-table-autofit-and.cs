using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableExample
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

            // First row, first cell – set a fixed width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Fixed width 100pt");

            // First row, second cell – set a different fixed width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Fixed width 200pt");

            // End the first row.
            builder.EndRow();

            // Second row, first cell – reuse the same width as the first column.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Another 100pt cell");

            // Second row, second cell – reuse the same width as the second column.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Another 200pt cell");

            // End the second row and the table.
            builder.EndRow();
            builder.EndTable();

            // Disable automatic resizing (AutoFit) and keep the column widths fixed.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);
            // Alternatively, you could set: table.AllowAutoFit = false;

            // Define the output file path.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");

            // Save the document.
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Optionally, inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
