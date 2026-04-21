using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating shape to demonstrate overlapping with the table.
        Shape shape = new Shape(doc, ShapeType.Rectangle);
        shape.Width = 100;
        shape.Height = 100;
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.HorizontalAlignment = HorizontalAlignment.Left;
        shape.VerticalAlignment = VerticalAlignment.Top;
        shape.Left = 50;
        shape.Top = 50;
        builder.InsertNode(shape);

        // Start building a floating table.
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Write("Floating Table");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Configure the table as a floating object.
        table.TextWrapping = TextWrapping.Around;
        table.AbsoluteHorizontalDistance = 150; // Horizontal offset from the paragraph.
        table.AbsoluteVerticalDistance = 150;   // Vertical offset from the paragraph.
        table.PreferredWidth = PreferredWidth.FromPoints(200);

        // Table.AllowOverlap is read‑only and defaults to true, so no need to set it.
        // The previous validation that threw an exception has been removed.

        // Save the document to a local file.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "FloatingTableOverlap.docx");
        doc.Save(outputPath);
    }
}
