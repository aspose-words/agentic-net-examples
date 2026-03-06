using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // <-- required for Table class

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with the required number of columns (example: 3 columns).
        Table table = builder.StartTable();

        // First row – column headers.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"Header {col + 1}");
        }
        builder.EndRow();

        // Second row – sample data.
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"Data {col + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as PDF.
        doc.Save("RenderedDocument.pdf", SaveFormat.Pdf);
    }
}
