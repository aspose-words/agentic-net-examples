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

        // First row, first cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Fixed width cell 1");

        // First row, second cell with a fixed width.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Fixed width cell 2");

        // End the first row.
        builder.EndRow();

        // Second row – cells inherit the column widths.
        builder.InsertCell();
        builder.Write("Row 2, cell 1");
        builder.InsertCell();
        builder.Write("Row 2, cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable automatic column resizing while preserving the existing widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to a local file.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableAutoFitDisabled.docx");
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
