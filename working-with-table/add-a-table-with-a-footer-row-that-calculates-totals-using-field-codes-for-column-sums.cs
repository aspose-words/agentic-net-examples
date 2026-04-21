using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class TableWithFooterExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Sample data rows.
        string[] items = { "Apples", "Bananas", "Carrots" };
        int[] quantities = { 20, 40, 50 };

        for (int i = 0; i < items.Length; i++)
        {
            builder.InsertCell();
            builder.Write(items[i]);
            builder.InsertCell();
            builder.Write(quantities[i].ToString());
            builder.EndRow();
        }

        // Footer row with a field that sums the values above.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a field code that calculates the sum of the numbers in the column above.
        builder.InsertField("=SUM(ABOVE)");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to a file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");

        // Optionally, inform that the process completed.
        Console.WriteLine($"Document saved successfully to: {outputPath}");
    }
}
