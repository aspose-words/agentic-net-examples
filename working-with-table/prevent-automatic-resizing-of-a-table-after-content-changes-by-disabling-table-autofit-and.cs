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

        // Start building the table.
        Table table = builder.StartTable();

        // First row – set fixed column widths.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Writeln("Header 1");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row – add long content that would normally trigger auto‑fit.
        builder.InsertCell();
        builder.Writeln("This is a very long piece of text that would normally cause the first column to expand if auto‑fit were enabled.");
        builder.InsertCell();
        builder.Writeln("Another long piece of text that would normally cause the second column to expand.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable automatic resizing and enforce fixed column widths.
        table.AllowAutoFit = false;
        table.AutoFit(AutoFitBehavior.FixedColumnWidths);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFixedWidth.docx");
        doc.Save(outputPath);

        // Simple verification that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
