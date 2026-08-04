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

        // Start a new table.
        Table table = builder.StartTable();

        // Define fixed column widths (in points). Example: three columns.
        double[] columnWidths = { 100, 150, 200 };

        // Build the first row and set the preferred width for each cell (column).
        for (int i = 0; i < columnWidths.Length; i++)
        {
            // Set the preferred width for the current column.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);

            // Insert the cell and add some sample text.
            builder.InsertCell();
            builder.Writeln($"Column {i + 1}");
        }
        // End the first row.
        builder.EndRow();

        // Add a second row with the same column widths.
        for (int i = 0; i < columnWidths.Length; i++)
        {
            // PreferredWidth must be set again for each new cell.
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(columnWidths[i]);

            builder.InsertCell();
            builder.Writeln($"Data {i + 1}");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Disable AutoFit to enforce the fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Optional: apply a built‑in table style for visual appearance.
        table.StyleIdentifier = StyleIdentifier.LightShadingAccent1;
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FixedColumnWidthsTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
