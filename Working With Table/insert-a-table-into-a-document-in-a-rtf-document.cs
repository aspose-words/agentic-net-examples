using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableIntoRtf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to simplify node insertion.
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
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Optionally set a title/description for accessibility.
        table.Title = "Sample Table";
        table.Description = "A simple 2x2 table inserted into an RTF document.";

        // Save the document as RTF.
        doc.Save("TableInRtf.rtf");
    }
}
