using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a header row.
        Table table = builder.StartTable();

        // Header cells.
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Description");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Sample data rows.
        AddRow(builder, "Apple", "Fresh red apples from the orchard.", "$1.20");
        AddRow(builder, "Banana", "Ripe bananas, sweet and soft.", "$0.80");
        AddRow(builder, "Cherry", "Organic cherries, packed in a box.", "$3.50");
        AddRow(builder, "Date", "Dry dates, high in fiber and nutrients.", "$2.10");
        AddRow(builder, "Elderberry", "Elderberries for making syrups and jams.", "$4.00");

        // End the table.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        string outputPath = "AdjustedTable.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }

    // Helper method to add a data row to the table.
    private static void AddRow(DocumentBuilder builder, string col1, string col2, string col3)
    {
        builder.InsertCell();
        builder.Write(col1);
        builder.InsertCell();
        builder.Write(col2);
        builder.InsertCell();
        builder.Write(col3);
        builder.EndRow();
    }
}
