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

        // Insert a table with a header row and three data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Number");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.EndRow();

        // Data rows.
        int[] numbers = { 10, 20, 30 };
        for (int i = 0; i < numbers.Length; i++)
        {
            // First column – start bookmark on the first data cell.
            builder.InsertCell();
            if (i == 0)
                builder.StartBookmark("ColNumbers");
            builder.Writeln(numbers[i].ToString());
            if (i == numbers.Length - 1)
                builder.EndBookmark("ColNumbers");

            // Second column.
            builder.InsertCell();
            builder.Writeln($"Item {i + 1}");
            builder.EndRow();
        }

        builder.EndTable();

        // Insert a paragraph after the table.
        builder.Writeln();

        // Insert a formula field that sums the bookmarked column.
        // Use the overload that takes a FieldType, then write the formula code.
        builder.InsertField(FieldType.FieldFormula, true);
        builder.Write("=SUM(ColNumbers)");
        builder.Writeln(); // Optional line break after the field result.

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
