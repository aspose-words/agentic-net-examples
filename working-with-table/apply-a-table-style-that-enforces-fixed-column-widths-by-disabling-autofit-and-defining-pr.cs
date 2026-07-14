using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // Define fixed column widths (in points).
        // 1 point = 1/72 inch.
        double[] columnWidths = { 100.0, 150.0, 200.0 };

        // Insert header row.
        for (int i = 0; i < columnWidths.Length; i++)
        {
            // Apply the preferred width to the current cell.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);
            builder.InsertCell();
            builder.Writeln($"Header {i + 1}");
        }
        builder.EndRow();

        // Insert a few data rows.
        for (int row = 0; row < 3; row++)
        {
            for (int col = 0; col < columnWidths.Length; col++)
            {
                // Reapply the same column width for each cell in this column.
                builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[col]);
                builder.InsertCell();
                builder.Writeln($"Row {row + 1}, Col {col + 1}");
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Disable AutoFit to enforce the fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FixedColumnWidthsTable.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
