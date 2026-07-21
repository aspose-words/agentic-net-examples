using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add a simple 2x2 grid.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a double border to each side of the table.
        // The last parameter 'true' overrides any explicit cell borders.
        table.SetBorder(BorderType.Left, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Right, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Top, LineStyle.Double, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Double, 1.5, Color.Black, true);

        // Save the document.
        doc.Save("TableDoubleBorder.docx");
    }
}
