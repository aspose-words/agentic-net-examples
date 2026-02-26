using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

class ExportGroupShapeAsPng
{
    static void Main()
    {
        // Load the source document (replace with your actual file path).
        Document doc = new Document("InputDocument.docx");

        // Find the first GroupShape in the document.
        // The GetChild method searches the whole document tree (deep = true).
        GroupShape groupShape = (GroupShape)doc.GetChild(NodeType.GroupShape, 0, true);
        if (groupShape == null)
        {
            Console.WriteLine("No group shape found in the document.");
            return;
        }

        // Create a renderer for the group shape.
        ShapeRenderer renderer = groupShape.GetShapeRenderer();

        // Define PNG image save options (optional – you can customize DPI, scaling, etc.).
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Render the group shape and save it as a PNG file.
        renderer.Save("GroupShape.png", pngOptions);

        Console.WriteLine("Group shape exported to GroupShape.png");
    }
}
