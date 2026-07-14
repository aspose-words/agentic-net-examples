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

        // First row, first cell with a fixed width of 100 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Writeln("Short text");

        // First row, second cell with a fixed width of 200 points.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Writeln("Longer text that might cause auto‑fit");

        // End the first row.
        builder.EndRow();

        // Second row, first cell with a long paragraph.
        builder.InsertCell();
        builder.Writeln("Very very very long text that would normally expand the column if auto‑fit were enabled.");

        // Second row, second cell with another long paragraph.
        builder.InsertCell();
        builder.Writeln("Another long text that could trigger auto‑fit.");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Disable automatic resizing and enforce fixed column widths.
        table.AllowAutoFit = false;
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to a local file.
        string outputPath = "TableFixedWidth.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");
    }
}
