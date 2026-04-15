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

        // Build a simple 2‑cell table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.EndTable();

        // Remove any existing borders from the table.
        table.ClearBorders();

        // Apply only top and bottom borders.
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        // Left and right borders remain cleared (no side borders).

        // Save the resulting document.
        doc.Save("TableWithTopBottomBorder.docx");
    }
}
