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

        // Start a table and add a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Product");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.InsertCell();
        builder.Write("Price");
        builder.EndRow();

        // Add some sample data rows.
        AddRow(builder, "Apples", "120", "$1.20");
        AddRow(builder, "Bananas", "85", "$0.80");
        AddRow(builder, "Cherries", "200", "$2.50");
        AddRow(builder, "Dates", "60", "$3.00");
        AddRow(builder, "Elderberries", "30", "$4.10");

        // End the table construction.
        builder.EndTable();

        // Adjust column widths proportionally to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "AdjustedTable.docx");
        doc.Save(outputPath);
    }

    // Helper method to insert a data row into the table.
    private static void AddRow(DocumentBuilder builder, string product, string quantity, string price)
    {
        builder.InsertCell();
        builder.Write(product);
        builder.InsertCell();
        builder.Write(quantity);
        builder.InsertCell();
        builder.Write(price);
        builder.EndRow();
    }
}
