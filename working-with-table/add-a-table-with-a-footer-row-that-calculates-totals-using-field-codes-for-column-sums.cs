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

        // ----- Header row -----
        builder.InsertCell();
        builder.Writeln("Item");
        builder.InsertCell();
        builder.Writeln("Quantity");
        builder.EndRow();

        // ----- Data rows -----
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
        builder.Writeln("40");
        builder.EndRow();

        // Row 3
        builder.InsertCell();
        builder.Writeln("Carrots");
        builder.InsertCell();
        builder.Writeln("50");
        builder.EndRow();

        // ----- Footer row with totals -----
        // First cell: label
        builder.InsertCell();
        builder.Writeln("Total");

        // Second cell: field that sums the values above in this column.
        // The field code "=SUM(ABOVE)" calculates the sum of numeric values in the column above.
        builder.InsertCell();
        builder.InsertField("=SUM(ABOVE)", "0");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file at '{outputPath}'.");
        }
    }
}
