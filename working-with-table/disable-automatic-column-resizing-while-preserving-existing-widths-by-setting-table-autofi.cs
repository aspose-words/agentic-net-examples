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

        // First cell with a fixed width of 100 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Cell 1 (100pt)");

        // Second cell with a fixed width of 150 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Cell 2 (150pt)");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Disable automatic column resizing while preserving the existing widths.
        // FixedColumnWidths disables auto‑fit behavior.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);
        // Also ensure the AllowAutoFit flag is turned off.
        table.AllowAutoFit = false;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAutoFitDisabled.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
