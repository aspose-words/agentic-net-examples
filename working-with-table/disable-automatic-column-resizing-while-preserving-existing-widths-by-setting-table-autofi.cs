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

        // First row – set explicit column widths.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Cell 1");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Write("Cell 2");

        builder.EndRow();

        // Second row – reuse the same column widths.
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Cell 3");

        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(200);
        builder.Write("Cell 4");

        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Disable automatic column resizing while preserving the set widths.
        table.AllowAutoFit = false;

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAutoFitDisabled.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved successfully.");

        // Optional confirmation (no user input required).
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
