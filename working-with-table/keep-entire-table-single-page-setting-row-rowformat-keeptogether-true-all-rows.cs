using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with several rows to demonstrate the setting.
        Table table = builder.StartTable();

        // Add a header row.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Add a few data rows.
        for (int i = 1; i <= 5; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Row {i} - Column 1");
            builder.InsertCell();
            builder.Writeln($"Row {i} - Column 2");
            builder.EndRow();
        }

        builder.EndTable();

        // Prevent each row from breaking across a page break.
        // This keeps the rows together, which effectively keeps the whole table on one page
        // as long as the table fits within the page height.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the modified document.
        doc.Save("Table.KeepTogether.docx");
    }
}
