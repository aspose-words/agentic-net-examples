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

        // Define fixed widths for the first column (100 points) and second column (200 points).
        // Disable AutoFit to enforce the fixed widths.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.InsertCell();
        builder.Writeln("Header 1");

        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add a second row with the same column widths.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.InsertCell();
        builder.Writeln("Row 1, Col 1");

        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.InsertCell();
        builder.Writeln("Row 1, Col 2");
        builder.EndRow();

        // Add a third row.
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.InsertCell();
        builder.Writeln("Row 2, Col 1");

        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.InsertCell();
        builder.Writeln("Row 2, Col 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Disable AutoFit to keep the column widths fixed.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to a local file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FixedWidthTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
