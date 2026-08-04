using System;
using System.IO;
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

        // Start a simple 2x2 table.
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

        // Apply borders: only top and bottom, no side borders.
        // First clear any existing borders.
        table.ClearBorders();

        // Top border.
        table.SetBorder(BorderType.Top, LineStyle.Single, 1.5, Color.Black, true);
        // Bottom border.
        table.SetBorder(BorderType.Bottom, LineStyle.Single, 1.5, Color.Black, true);
        // Left and right borders are left cleared (no border).

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableStyleTopBottom.docx");
        doc.Save(outputPath);
    }
}
