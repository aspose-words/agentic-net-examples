using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFixedLayout
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

            // ---- First Row ----
            // First cell – set explicit width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Column 1, Row 1");

            // Second cell – set explicit width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Column 2, Row 1");

            // End the first row.
            builder.EndRow();

            // ---- Second Row ----
            // First cell – reuse the same width as column 1.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Column 1, Row 2");

            // Second cell – reuse the same width as column 2.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Column 2, Row 2");

            // End the second row.
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Disable AutoFit and enforce fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Ensure the output directory exists.
            string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
            Directory.CreateDirectory(outputDir);

            // Save the document.
            string outputPath = Path.Combine(outputDir, "FixedLayoutTable.docx");
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The document was not saved correctly.");

            // Optional: inform that the process completed.
            Console.WriteLine($"Document saved to: {outputPath}");
        }
    }
}
