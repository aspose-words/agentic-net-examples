using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two overlapping floating shapes.
        // First shape – rectangle.
        builder.InsertShape(ShapeType.Rectangle, 150, 100);
        Shape rectShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];
        rectShape.WrapType = WrapType.None; // Make it floating.

        // Second shape – ellipse (will be on top initially).
        builder.InsertShape(ShapeType.Ellipse, 150, 100);
        Shape ellipseShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[1];
        ellipseShape.WrapType = WrapType.None; // Make it floating.

        // Retrieve all shapes in the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        if (shapeNodes.Count < 2)
            throw new Exception("Expected at least two shapes in the document.");

        // Cast the first shape (rectangle).
        Shape firstShape = (Shape)shapeNodes[0];

        // Get its initial Z‑order index.
        int initialZOrder = firstShape.ZOrder;

        // Bring the first shape to the front by moving it to the end of its parent's child collection.
        CompositeNode parent = (CompositeNode)firstShape.ParentNode;
        parent.RemoveChild(firstShape);
        parent.AppendChild(firstShape);

        // Get its new Z‑order index.
        int newZOrder = firstShape.ZOrder;

        // Save the document.
        string outputPath = "ZOrderExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");

        // Output the Z‑order values (optional verification).
        Console.WriteLine($"Initial Z‑order: {initialZOrder}");
        Console.WriteLine($"New Z‑order after bringing to front: {newZOrder}");
    }
}
