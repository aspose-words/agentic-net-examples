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

        // Header row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Add a few data rows.
        for (int i = 1; i <= 4; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i} Col 1");
            builder.InsertCell();
            builder.Write($"Row {i} Col 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable banded rows and columns.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithBanding.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Indicate success.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
