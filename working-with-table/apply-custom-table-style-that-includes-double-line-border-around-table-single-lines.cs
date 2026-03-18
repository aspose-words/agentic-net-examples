using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a simple 2x2 grid.
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

        // Finish table construction.
        builder.EndTable();

        // Remove any existing borders.
        table.ClearBorders();

        // Apply a double line border to the outer edges of the table.
        table.SetBorder(BorderType.Left,   LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Right,  LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Top,    LineStyle.Double, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 2.0, Color.Black, true);

        // Apply single line borders to the internal grid (horizontal and vertical lines).
        table.SetBorder(BorderType.Horizontal, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Vertical,   LineStyle.Single, 1.0, Color.Black, true);

        // Save the document.
        doc.Save("CustomTableStyle.docx");
    }
}
