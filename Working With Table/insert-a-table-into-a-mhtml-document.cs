using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class InsertTableIntoMhtml
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");
        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Value 1");
        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Value 2");
        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optionally apply auto‑fit to contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as MHTML.
        doc.Save("TableDocument.mhtml", SaveFormat.Mhtml);
    }
}
