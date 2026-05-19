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

            // Start a table.
            Table table = builder.StartTable();

            // Define fixed widths for three columns (in points).
            double[] columnWidths = { 100, 150, 200 };

            // First row – set widths and add header text.
            for (int i = 0; i < columnWidths.Length; i++)
            {
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);
                builder.InsertCell();
                builder.Writeln($"Header {i + 1}");
            }
            builder.EndRow();

            // Second row – reuse the same column widths.
            for (int i = 0; i < columnWidths.Length; i++)
            {
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);
                builder.InsertCell();
                builder.Writeln($"Row 1, Col {i + 1}");
            }
            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Disable AutoFit to enforce the fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FixedColumnWidthsTable.docx");
            doc.Save(outputPath);

            // Verify that the file was saved.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("Document was not saved correctly.");
        }
    }
}
