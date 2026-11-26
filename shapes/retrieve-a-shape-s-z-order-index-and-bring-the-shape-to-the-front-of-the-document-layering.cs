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

        // Insert first shape (rectangle) and keep a reference.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.WrapType = WrapType.None; // Ensure shape is floating.
        rect.Left = 0;
        rect.Top = 0;
        rect.StrokeColor = System.Drawing.Color.Blue;
        rect.FillColor = System.Drawing.Color.LightBlue;

        // Insert second shape (ellipse) that overlaps the first one.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 100);
        ellipse.WrapType = WrapType.None;
        ellipse.Left = 50; // Overlap with rectangle.
        ellipse.Top = 30;
        ellipse.StrokeColor = System.Drawing.Color.Red;
        ellipse.FillColor = System.Drawing.Color.Pink;

        // Shapes are inserted into the same paragraph. Retrieve their positions within that paragraph.
        Paragraph parentParagraph = (Paragraph)rect.ParentNode;

        // Get the collection of shape nodes within the paragraph.
        NodeCollection shapeNodes = parentParagraph.GetChildNodes(NodeType.Shape, true);
        int initialRectIndex = shapeNodes.IndexOf(rect);
        int initialEllipseIndex = shapeNodes.IndexOf(ellipse);

        // Bring the rectangle to the front by moving it after the ellipse in the node collection.
        rect.Remove(); // Detach from its current position.
        parentParagraph.InsertAfter(rect, ellipse); // Insert after the ellipse.

        // Verify that the rectangle now has a higher index (i.e., is on top).
        NodeCollection updatedShapeNodes = parentParagraph.GetChildNodes(NodeType.Shape, true);
        int finalRectIndex = updatedShapeNodes.IndexOf(rect);
        int finalEllipseIndex = updatedShapeNodes.IndexOf(ellipse);

        if (finalRectIndex <= finalEllipseIndex)
        {
            throw new InvalidOperationException("Failed to bring the shape to the front.");
        }

        // Save the document.
        string outputPath = "ZOrderExample.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not created.", outputPath);
        }

        // Output result to console (no user interaction required).
        Console.WriteLine($"Initial rectangle index: {initialRectIndex}, Final rectangle index: {finalRectIndex}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
