using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("Input.doc");

        // Retrieve all Shape nodes (including text boxes) from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removing a shape does not affect the loop index.
        for (int i = shapeNodes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapeNodes[i];

            // Identify text box shapes by their ShapeType.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the text box from the document.
                shape.Remove();
            }
        }

        // Prepare image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Save the modified document as a multi‑page TIFF file.
        doc.Save("Output.tiff", tiffOptions);
    }
}
