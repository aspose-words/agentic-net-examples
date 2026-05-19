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

        // Build a simple 2x2 table.
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
        builder.EndTable();

        // Remove any existing borders.
        table.ClearBorders();

        // Apply only top and bottom borders.
        table.SetBorder(BorderType.Top, LineStyle.Single, 2.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 2.0, Color.Black, true);
        // Left and right borders remain cleared (no side borders).

        // Save the document to the local file system.
        string outputPath = "TableStyleTopBottom.docx";
        doc.Save(outputPath);
    }
}
