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
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        // Set a fixed width for the cell (column).
        builder.CurrentParagraph.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);
        builder.Writeln("Header 1");

        // First row, second cell.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
        builder.Writeln("Header 2");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(80);
        builder.Writeln("Data 1");

        // Second row, second cell.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(120);
        builder.Writeln("Data 2");

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Disable auto‑fit so column widths remain fixed.
        table.AllowAutoFit = false;

        // Save the document.
        string outputPath = "TableAutoFitDisabled.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("Document saved to " + Path.GetFullPath(outputPath));
    }
}
