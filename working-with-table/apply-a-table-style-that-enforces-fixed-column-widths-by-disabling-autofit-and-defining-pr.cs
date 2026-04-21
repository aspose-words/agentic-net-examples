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

            // ----- Header row -----
            // First column – 100 points width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Header 1");

            // Second column – 150 points width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Header 2");

            // Third column – 200 points width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Header 3");

            builder.EndRow();

            // ----- Data row -----
            // First column – same width as header.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Row 1, Col 1");

            // Second column – same width as header.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Row 1, Col 2");

            // Third column – same width as header.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Row 1, Col 3");

            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Disable AutoFit to enforce the fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "FixedColumnWidthsTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output document was not saved correctly.");

            // Optional: inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
