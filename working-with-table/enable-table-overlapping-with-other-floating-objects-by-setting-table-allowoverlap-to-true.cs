using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape shape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        shape.WrapType = WrapType.Square;
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        shape.Left = 0;   // Position at the left edge of the page.
        shape.Top = 0;    // Position at the top edge of the page.
        shape.AllowOverlap = true; // Allow overlapping with other floating objects.

        // Move the builder below the shape to start the table.
        builder.Writeln(); // Ensure we are in a new paragraph.

        // Build a simple 2x2 table.
        builder.StartTable();
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

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // In the current Aspose.Words version, Table does not expose
        // AllowOverlap, WrapType, RelativeHorizontalPosition, etc.
        // Overlapping with the shape is allowed by default, so no additional
        // properties need to be set on the table.

        // Save the document.
        string outputPath = "TableOverlapExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }
    }
}
