using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with three rows and two columns.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Value");
        builder.EndRow();

        // Data rows with numeric values in the second column.
        int[] numbers = { 10, 20, 30 };
        for (int i = 0; i < numbers.Length; i++)
        {
            // First column: item label.
            builder.InsertCell();
            builder.Write($"Item {i + 1}");

            // Second column: numeric value with a bookmark.
            builder.InsertCell();
            string bookmarkName = $"num{i + 1}";
            builder.StartBookmark(bookmarkName);
            builder.Write(numbers[i].ToString());
            builder.EndBookmark(bookmarkName);

            builder.EndRow();
        }

        builder.EndTable();

        // Insert a paragraph after the table.
        builder.Writeln();

        // Insert a formula field that sums the bookmarked numbers.
        // The field code will be: =SUM(num1,num2,num3)
        string fieldCode = "=SUM(num1,num2,num3)";
        builder.InsertField(fieldCode);

        // Update fields so the sum is calculated.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception("Failed to create the output document.");
        }
    }
}
