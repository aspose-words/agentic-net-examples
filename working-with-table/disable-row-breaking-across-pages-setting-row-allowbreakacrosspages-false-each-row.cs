using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with enough rows to span multiple pages.
        Table table = builder.StartTable();

        // Add a header row.
        builder.InsertCell();
        builder.Writeln("Header");
        builder.EndRow();

        // Add many rows to force the table to span at least two pages.
        for (int i = 1; i <= 50; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i}");
            builder.EndRow();
        }

        builder.EndTable();

        // Disable the ability for each row to break across a page boundary.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the updated document.
        doc.Save("Table.AllowBreakAcrossPages.docx");
        Console.WriteLine("Document saved successfully.");
    }
}
