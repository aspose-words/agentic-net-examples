using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑column table.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Column 1");
        builder.InsertCell();
        builder.Writeln("Column 2");
        builder.InsertCell();
        builder.Writeln("Column 3");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.InsertCell();
        builder.Writeln("Data 3");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Disable AutoFit to enforce fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Define preferred widths for each column (using points; 1 inch = 72 points).
        if (table.Rows.Count > 0 && table.Rows[0].Cells.Count >= 3)
        {
            // First column – 2 inches.
            table.Rows[0].Cells[0].CellFormat.PreferredWidth =
                PreferredWidth.FromPoints(2 * 72);

            // Second column – 1.5 inches.
            table.Rows[0].Cells[1].CellFormat.PreferredWidth =
                PreferredWidth.FromPoints(1.5 * 72);

            // Third column – 2.5 inches.
            table.Rows[0].Cells[2].CellFormat.PreferredWidth =
                PreferredWidth.FromPoints(2.5 * 72);
        }

        // Save the document.
        string outputPath = "FixedWidthTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Confirmation (no user interaction required).
        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
