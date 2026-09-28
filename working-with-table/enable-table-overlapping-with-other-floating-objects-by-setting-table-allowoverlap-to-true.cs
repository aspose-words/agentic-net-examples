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

        // Insert a floating rectangle shape that will overlap the table.
        builder.InsertShape(ShapeType.Rectangle, 100, 50);
        // Retrieve the inserted shape.
        Shape shape = (Shape)builder.CurrentParagraph.GetChildNodes(NodeType.Shape, true)[0];
        // Ensure the shape does not wrap text so it can overlap other objects.
        shape.WrapType = WrapType.None;

        // Move to a new paragraph to start the table.
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
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Overlapping is allowed by default when surrounding objects have no text wrapping.
        // No need to set Table.AllowOverlap because it is read‑only.

        // Save the document.
        string outputPath = "TableOverlap.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
