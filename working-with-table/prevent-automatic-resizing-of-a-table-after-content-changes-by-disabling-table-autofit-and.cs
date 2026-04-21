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

        // Start a table with three columns.
        Table table = builder.StartTable();

        // Set fixed column widths (in points) for each cell in the first row.
        // These widths will be applied to all cells in the respective columns.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Writeln("Header 1");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Writeln("Header 2");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Writeln("Header 3");

        builder.EndRow();

        // Add a second row with normal content.
        builder.InsertCell();
        builder.Writeln("Short text");

        builder.InsertCell();
        builder.Writeln("More text");

        builder.InsertCell();
        builder.Writeln("Even more text");

        builder.EndRow();

        // Add a third row with a very long text to demonstrate that the table will NOT auto‑fit.
        builder.InsertCell();
        builder.Writeln("This is a very long piece of text that would normally cause the column to expand if auto‑fit were enabled. " +
                         "Because we have disabled auto‑fit and set fixed column widths, the text will be truncated or wrapped according to the cell's settings.");

        builder.InsertCell();
        builder.Writeln("Another long text that should stay within the fixed width of its column.");

        builder.InsertCell();
        builder.Writeln("Yet another long text to test the fixed column behavior.");

        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Disable automatic resizing (auto‑fit) for the entire table.
        table.AllowAutoFit = false;
        // Alternatively, you can call AutoFit with FixedColumnWidths behavior.
        // table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OutputTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");
        }

        // Inform the user (optional, not required for non‑interactive execution).
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
