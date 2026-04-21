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

        // Start a table and add a few rows with two columns each.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        string[,] data = { { "Apples", "20" }, { "Bananas", "40" }, { "Carrots", "50" } };
        for (int i = 0; i < data.GetLength(0); i++)
        {
            builder.InsertCell();
            builder.Write(data[i, 0]);
            builder.InsertCell();
            builder.Write(data[i, 1]);
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable both row banding and column banding.
        table.StyleOptions = TableStyleOptions.RowBands | TableStyleOptions.ColumnBands;

        // Optional: Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithBandedRowsAndColumns.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        // Inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
