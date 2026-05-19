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

        // Start building a table.
        Table table = builder.StartTable();

        // First cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("First column fixed width.");

        // Second cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Second column fixed width.");

        // End the first row.
        builder.EndRow();

        // Add a second row (optional).
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable automatic table auto‑fit to keep column widths fixed.
        table.AllowAutoFit = false;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableAllowAutoFitFalse.docx");
        doc.Save(outputPath);
    }
}
