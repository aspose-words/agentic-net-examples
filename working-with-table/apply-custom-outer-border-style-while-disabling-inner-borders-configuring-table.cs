using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Remove all existing borders (both table and cell borders).
        table.ClearBorders();

        // Apply a custom outer border to the table.
        // The last argument (true) overrides any explicit cell borders, effectively disabling inner borders.
        table.SetBorder(BorderType.Left,   LineStyle.Double, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 2.0, Color.Blue, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Blue, true);

        // Save the document to disk.
        doc.Save("TableCustomOuterBorder.docx");
    }
}
