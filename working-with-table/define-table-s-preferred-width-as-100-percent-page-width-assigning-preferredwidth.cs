using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetTablePreferredWidth
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table and add a single row with three cells as an example.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell #1");
        builder.InsertCell();
        builder.Write("Cell #2");
        builder.InsertCell();
        builder.Write("Cell #3");
        builder.EndRow();
        builder.EndTable();

        // Set the table's preferred width to 100 % of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Save the document to disk.
        doc.Save("TableWithFullWidth.docx");
    }
}
