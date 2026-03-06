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

        // Use DocumentBuilder to simplify inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row – two cells.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row – two cells.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndTable(); // Ends the table.

        // Optional: set a title/description for accessibility.
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into an MHTML document.";

        // Save the document as MHTML (MHT) format.
        // SaveFormat.Mhtml ensures the output is a single MHTML file.
        doc.Save("TableInMhtml.mht", SaveFormat.Mhtml);
    }
}
