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
        builder.StartTable();

        // First row - three cells with fixed widths.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            // Set a fixed width for the cell (e.g., 100 points).
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln($"Row 1, Cell {i + 1}");
        }
        builder.EndRow();

        // Second row - same column widths.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
            builder.Writeln($"Row 2, Cell {i + 1}");
        }
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table and disable auto‑fit.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        table.AllowAutoFit = false;

        // Save the document.
        string outputPath = "TableAutoFitDisabled.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
