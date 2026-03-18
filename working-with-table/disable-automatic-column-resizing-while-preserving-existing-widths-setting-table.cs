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

        // Start building a table.
        Table table = builder.StartTable();

        // ---- First row ----
        // First cell – set an explicit width (100 points).
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(100);
        builder.Write("Cell 1");

        // Second cell – set an explicit width (150 points).
        builder.InsertCell();
        builder.CellFormat.PreferredWidth = PreferredWidth.FromPoints(150);
        builder.Write("Cell 2");
        builder.EndRow();

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Cell 3");

        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Disable automatic column resizing.
        // This preserves the widths we set above.
        table.AllowAutoFit = false;

        // Save the document to disk.
        doc.Save("Table_NoAutoFit.docx");
    }
}
