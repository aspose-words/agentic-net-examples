using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3‑row, 2‑column table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();

        // Data rows.
        builder.InsertCell();
        builder.Write("Apples");
        builder.InsertCell();
        builder.Write("10");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Bananas");
        builder.InsertCell();
        builder.Write("20");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Carrots");
        builder.InsertCell();
        builder.Write("30");
        builder.EndRow();

        // Add a row that will contain the sum of the "Quantity" column.
        builder.InsertCell();
        builder.Write("Total");
        builder.InsertCell();

        // Insert a formula field that sums the numbers above in this column.
        // The field code is: = SUM(ABOVE)
        // Using the string overload automatically updates the field result.
        builder.InsertField("= SUM(ABOVE) \\* MERGEFORMAT");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Insert a paragraph after the table that also shows the sum using a field.
        builder.Writeln();
        builder.Write("Sum of quantities (outside table): ");
        // The same formula works when placed in a paragraph after the table,
        // because the field evaluates the numbers in the preceding table column.
        builder.InsertField("= SUM(ABOVE) \\* MERGEFORMAT");
        builder.Writeln();

        // Ensure all fields are up‑to‑date.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TableColumnSum.docx");
    }
}
