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

        // Start the table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // ---------- Data rows ----------
        // Row 1
        builder.InsertCell();
        builder.Writeln("Apples");
        builder.InsertCell();
        builder.Writeln("20");
        builder.EndRow();

        // Row 2
        builder.InsertCell();
        builder.Writeln("Bananas");
        builder.InsertCell();
        builder.Writeln("35");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("Carrots");
        builder.InsertCell();
        builder.Writeln("15");
        builder.EndRow();

        // ---------- Footer row with totals ----------
        builder.InsertCell();
        builder.Writeln("Total");

        // Insert a field that sums all numbers above in the current column.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", null);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // Optionally, inform that the process completed.
        Console.WriteLine("Document created successfully at: " + outputPath);
    }
}
