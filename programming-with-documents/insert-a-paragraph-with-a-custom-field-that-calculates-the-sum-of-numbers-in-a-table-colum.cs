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

        // Build a simple 2x2 table with numeric values in the first column.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("10");          // Numeric value.
        builder.InsertCell();
        builder.Write("Item A");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("20");          // Numeric value.
        builder.InsertCell();
        builder.Write("Item B");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Insert a new paragraph after the table.
        builder.Writeln(); // Paragraph break.
        builder.Write("Sum of the first column: ");

        // Insert a formula field that calculates the sum of the numbers above it.
        // The field code is inserted without the surrounding braces.
        builder.InsertField("= SUM(ABOVE) ");

        // Update all fields so the result is calculated.
        doc.UpdateFields();

        // Save the document to the local file system.
        string outputPath = "SumTableField.docx";
        doc.Save(outputPath);
    }
}
