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

        // Start a table with two columns.
        builder.StartTable();

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

        // Row that will contain the sum of the numbers in the second column.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();
        // Insert a formula field that sums the cells above in this column.
        // The field is placed inside a paragraph (the cell's paragraph).
        builder.InsertField("= SUM(ABOVE) ");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Update all fields so the sum is calculated.
        doc.UpdateFields();

        // Save the document.
        doc.Save("SumField.docx");
    }
}
