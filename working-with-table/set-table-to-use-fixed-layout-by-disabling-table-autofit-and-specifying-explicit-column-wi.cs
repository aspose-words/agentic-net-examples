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

            // First row, first cell – set explicit width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Column 1");

            // First row, second cell – set explicit width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Column 2");

            // First row, third cell – set explicit width.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Column 3");

            // End the first row.
            builder.EndRow();

            // Add a second row with the same column widths.
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln("Data 1");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
            builder.Writeln("Data 2");

            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
            builder.Writeln("Data 3");

            builder.EndRow();

            // Finish the table.
            builder.EndTable();

            // Disable AutoFit to enforce fixed column widths.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document.
            string outputPath = "FixedLayoutTable.docx";
            doc.Save(outputPath);

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
