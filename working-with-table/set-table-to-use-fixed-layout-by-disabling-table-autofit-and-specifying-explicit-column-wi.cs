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

        // ---- First row (header) ----
        // Cell 1
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Writeln("Header 1");

        // Cell 2
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Writeln("Header 2");

        // Cell 3
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Writeln("Header 3");

        // End the first row.
        builder.EndRow();

        // ---- Second row (data) ----
        // Cell 1
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Writeln("Data 1");

        // Cell 2
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Writeln("Data 2");

        // Cell 3
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Writeln("Data 3");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable AutoFit to enforce fixed column widths.
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FixedLayoutTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
