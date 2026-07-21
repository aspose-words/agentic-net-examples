using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row table with numeric values in the first column.
        builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Data rows.
        int[] numbers = { 10, 20, 30 };
        foreach (int num in numbers)
        {
            builder.InsertCell();
            builder.Write($"Number {num}");
            builder.InsertCell();
            builder.Write(num.ToString());
            builder.EndRow();
        }

        // Row that will contain the sum formula in the second column.
        builder.InsertCell();
        builder.Write("Sum");
        builder.InsertCell();

        // Start a bookmark that will surround the formula field.
        builder.StartBookmark("SumCell");

        // Insert a formula field that sums the cells above in this column.
        // Use the overload that accepts a field code string.
        builder.InsertField("= SUM(ABOVE) ");

        // End the bookmark.
        builder.EndBookmark("SumCell");

        // End the table.
        builder.EndTable();

        // Insert a new paragraph after the table.
        builder.Writeln();

        // Write a description and insert a REF field that displays the bookmarked sum.
        builder.Write("Total sum of the column: ");
        // Insert a REF field that references the bookmark "SumCell".
        builder.InsertField("REF SumCell");

        // Update all fields in the document so that the formula and REF fields show correct results.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
