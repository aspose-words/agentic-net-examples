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

        // Insert a floating shape that will be positioned near the table.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Use WrapType.None to make the shape floating.
        shape.WrapType = WrapType.None;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.HorizontalAlignment = HorizontalAlignment.Right;
        shape.VerticalAlignment = VerticalAlignment.Top;
        shape.Top = 0;
        shape.Left = 0;

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Write("Floating Table");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Configure the table to wrap text around it.
        table.TextWrapping = TextWrapping.Around;

        // Table.AllowOverlap is read‑only and defaults to true, so no assignment is needed.

        // Position the floating table.
        table.HorizontalAnchor = RelativeHorizontalPosition.Page;
        table.VerticalAnchor = RelativeVerticalPosition.Page;
        table.AbsoluteHorizontalDistance = 50; // Move 50 points from the anchor.
        table.AbsoluteVerticalDistance = 50;   // Move 50 points from the anchor.

        // Save the document to a local file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAllowOverlap.docx");
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not saved correctly.");
    }
}
