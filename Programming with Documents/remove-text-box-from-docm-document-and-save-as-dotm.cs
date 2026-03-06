using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("Input.docm");

        // Iterate through all Shape nodes in the document.
        int shapeIndex = 0;
        Shape shape = doc.GetChild(NodeType.Shape, shapeIndex, true) as Shape;
        while (shape != null)
        {
            // If the shape is a text box, remove it from its parent.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();

            shapeIndex++;
            shape = doc.GetChild(NodeType.Shape, shapeIndex, true) as Shape;
        }

        // Save the modified document as a DOTM template.
        doc.Save("Output.dotm", SaveFormat.Dotm);
    }
}
