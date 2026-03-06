using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Find and remove every shape that is a text box.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Prepare TIFF save options (render all pages into a multi‑frame TIFF).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // By default all pages are rendered; the following line explicitly sets that behavior.
        tiffOptions.PageSet = new PageSet(0);

        // Save the modified document as a TIFF image.
        doc.Save("output.tiff", tiffOptions);
    }
}
