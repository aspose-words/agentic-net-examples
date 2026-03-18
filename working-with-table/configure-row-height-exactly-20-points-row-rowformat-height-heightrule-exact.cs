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

        // Begin a table.
        Table table = builder.StartTable();

        // First row – uses default formatting.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Configure the next row to have an exact height of 20 points.
        builder.RowFormat.Height = 20;                 // Height in points.
        builder.RowFormat.HeightRule = HeightRule.Exactly; // Exact rule.

        // Add cells for the second row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("RowHeightExact20.docx");
    }
}
