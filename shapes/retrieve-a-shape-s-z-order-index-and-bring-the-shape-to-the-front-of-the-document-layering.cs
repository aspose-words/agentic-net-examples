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

        // Insert three overlapping rectangle shapes.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        shape1.Name = "Shape1";
        shape1.WrapType = WrapType.None;
        shape1.Left = 100;
        shape1.Top = 100;

        Shape shape2 = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        shape2.Name = "TargetShape";
        shape2.WrapType = WrapType.None;
        shape2.Left = 120; // Overlap with shape1
        shape2.Top = 120;

        Shape shape3 = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        shape3.Name = "Shape3";
        shape3.WrapType = WrapType.None;
        shape3.Left = 140; // Overlap with previous shapes
        shape3.Top = 140;

        // Retrieve the target shape via traversal.
        Shape target = null;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape s in shapes)
        {
            if (s.Name == "TargetShape")
            {
                target = s;
                break;
            }
        }

        if (target == null)
            throw new InvalidOperationException("Target shape not found.");

        // Get the current Z‑order index of the target shape.
        int originalZOrder = target.ZOrder;
        Console.WriteLine($"Original Z‑order index of target shape: {originalZOrder}");

        // Bring the target shape to the front of the layering.
        // Move the shape node to the end of its parent composite node's child collection.
        Node parent = target.ParentNode;
        if (parent is CompositeNode compositeParent)
        {
            target.Remove();               // Detach from current position.
            compositeParent.AppendChild(target); // Re‑append as the last child (frontmost).
        }
        else
        {
            throw new InvalidOperationException("Unexpected parent node type for the shape.");
        }

        // Verify that the Z‑order index has changed (it should now be the highest).
        int newZOrder = target.ZOrder;
        Console.WriteLine($"New Z‑order index of target shape after BringToFront: {newZOrder}");

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
