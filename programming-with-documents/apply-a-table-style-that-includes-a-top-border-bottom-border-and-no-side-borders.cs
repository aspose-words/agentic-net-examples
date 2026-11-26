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

        // Apply borders: visible top and bottom, no side borders.
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, Color.Black, true);
        table.SetBorder(BorderType.Left, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Right, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Horizontal, LineStyle.None, 0, Color.Empty, true);
        table.SetBorder(BorderType.Vertical, LineStyle.None, 0, Color.Empty, true);

        // Save the document to disk.
        string outputPath = "TableStyle.docx";
        doc.Save(outputPath);
    }
}
