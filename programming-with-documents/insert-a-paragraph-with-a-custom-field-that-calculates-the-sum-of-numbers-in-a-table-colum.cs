using System;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a simple 3‑row, 2‑column table with numeric values in the first column.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Write("Item");
            builder.InsertCell();
            builder.Write("Description");
            builder.EndRow();

            // First data row.
            builder.InsertCell();
            builder.Write("10");
            builder.InsertCell();
            builder.Write("Apples");
            builder.EndRow();

            // Second data row.
            builder.InsertCell();
            builder.Write("20");
            builder.InsertCell();
            builder.Write("Bananas");
            builder.EndRow();

            // End the table.
            builder.EndTable();

            // Move the cursor to a new paragraph after the table.
            builder.Writeln();
            builder.Write("Sum of the first column: ");

            // Insert a formula field that sums the numbers above it.
            // The field is placed in the paragraph (outside the table) but still works because
            // Word evaluates the SUM(ABOVE) based on the nearest table column above the field.
            builder.InsertField("=SUM(ABOVE) ");

            // Update all fields so the result is calculated.
            doc.UpdateFields();

            // Save the document to the local file system.
            doc.Save("Output.docx");
        }
    }
}
