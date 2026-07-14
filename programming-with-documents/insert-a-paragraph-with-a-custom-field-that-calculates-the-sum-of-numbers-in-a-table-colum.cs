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

        // Build a simple table with a header and three data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // First data row.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        // Third data row.
        builder.InsertCell();
        builder.Write("Carrots");
        builder.InsertCell();
        builder.Write("30");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Insert a new paragraph after the table.
        builder.Writeln();

        // Write a label for the total.
        builder.Write("Total quantity: ");

        // Insert a formula field that sums the numbers above in the same column.
        // The field code "=SUM(ABOVE)" tells Word to add all numeric values in the column above this cell.
        // Use the overload that does not update immediately; we will update all fields later.
        builder.InsertField("=SUM(ABOVE)", "");

        // Recalculate all fields so the SUM field shows the correct result.
        doc.UpdateFields();

        // Save the document to disk.
        doc.Save("TableSumField.docx");
    }
}
