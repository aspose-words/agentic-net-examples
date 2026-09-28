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

        // Insert a floating shape (rectangle) that will overlap the table.
        Shape floatingShape = builder.InsertShape(ShapeType.Rectangle, 100, 50);
        floatingShape.WrapType = WrapType.Square;
        floatingShape.VerticalAlignment = VerticalAlignment.Top;
        floatingShape.HorizontalAlignment = HorizontalAlignment.Left;
        floatingShape.Left = 0;
        floatingShape.Top = 0;

        // Move the builder after the shape to start the table.
        builder.Writeln(); // Ensure we are in a new paragraph.

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // The Table.AllowOverlap property is read‑only in this version of Aspose.Words.
        // Overlapping with floating objects is enabled by default, so no explicit assignment is required.

        // Save the document.
        string outputPath = "TableOverlap.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");
    }
}
