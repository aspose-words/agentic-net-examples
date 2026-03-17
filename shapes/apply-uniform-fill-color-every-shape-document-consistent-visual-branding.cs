using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables; // Required for NodeType enum

class ApplyUniformShapeFill
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Retrieve all shapes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Apply a solid fill color to each shape.
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Set the fill to a uniform solid color (e.g., LightSteelBlue).
            shape.Fill.Solid(Color.LightSteelBlue);
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
