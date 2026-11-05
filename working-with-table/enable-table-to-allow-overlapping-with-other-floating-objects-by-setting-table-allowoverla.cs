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

        // Insert a floating shape (a rectangle) to demonstrate overlapping.
        Shape floatingShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        floatingShape.WrapType = WrapType.Square;
        floatingShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        floatingShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        floatingShape.HorizontalAlignment = HorizontalAlignment.Right;
        floatingShape.VerticalAlignment = VerticalAlignment.Top;

        // Move the builder to a new paragraph after the shape.
        builder.Writeln();

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

        // The Table.AllowOverlap property is read‑only in this version of Aspose.Words.
        // Overlapping with floating objects is enabled by default for tables,
        // so no additional code is required here.

        // Save the document.
        string outputPath = "TableAllowOverlap.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
