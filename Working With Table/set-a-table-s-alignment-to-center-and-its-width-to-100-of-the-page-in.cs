using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. At least one row and cell are required.
        Table table = builder.StartTable();

        // Insert a single cell with placeholder text.
        builder.InsertCell();
        builder.Write("Sample text");

        // Finish the table.
        builder.EndTable();

        // Center the table on the page.
        table.Alignment = TableAlignment.Center;

        // Set the table width to 100% of the page width.
        table.PreferredWidth = PreferredWidth.FromPercent(100);

        // Save the document in WORDML (XML) format.
        doc.Save("TableCentered100Percent.xml", SaveFormat.WordML);
    }
}
