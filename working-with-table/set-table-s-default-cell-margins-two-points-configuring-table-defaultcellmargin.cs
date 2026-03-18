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

        // Set the default cell padding (margins) for all cells that will be created.
        // The values are in points.
        builder.CellFormat.TopPadding = 2;
        builder.CellFormat.BottomPadding = 2;
        builder.CellFormat.LeftPadding = 2;
        builder.CellFormat.RightPadding = 2;

        // Build a simple 2‑column table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("TableDefaultCellMargin.docx");
    }
}
