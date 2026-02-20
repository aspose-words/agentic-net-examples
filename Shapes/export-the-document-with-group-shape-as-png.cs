using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class ExportGroupShapeAsPng
{
    static void Main()
    {
        // Load the source document that contains a group shape.
        Document doc = new Document("GroupShapeDocument.docx");

        // Find the first group shape in the document.
        // A group shape is a Shape whose IsGroup property is true.
        Shape groupShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
        if (groupShape == null || !groupShape.IsGroup)
            throw new InvalidOperationException("No group shape found in the document.");

        // Create a ShapeRenderer for the group shape.
        ShapeRenderer renderer = new ShapeRenderer(groupShape);

        // Define image save options for PNG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

        // Render the group shape to a PNG file.
        renderer.Save("GroupShape.png", saveOptions);
    }
}
