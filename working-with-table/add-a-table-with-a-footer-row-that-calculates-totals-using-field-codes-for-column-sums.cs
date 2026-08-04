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
        builder.Font.Bold = true;
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // ----- Data rows -----
        AddDataRow(builder, "Apples", "20");
        AddDataRow(builder, "Bananas", "40");
        AddDataRow(builder, "Carrots", "50");

        // ----- Footer row with totals -----
        builder.InsertCell();
        builder.Font.Bold = true;
        builder.Write("Total");
        builder.InsertCell();

        // Insert a field that sums the values above in the same column.
        // The field code "=SUM(ABOVE)" calculates the sum of numeric values in the column.
        builder.InsertField("=SUM(ABOVE)", null);
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableWithFooter.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }

    // Helper method to add a data row with two cells.
    private static void AddDataRow(DocumentBuilder builder, string item, string quantity)
    {
        builder.InsertCell();
        builder.Font.Bold = false;
        builder.Write(item);
        builder.InsertCell();
        builder.Write(quantity);
        builder.EndRow();
    }
}
