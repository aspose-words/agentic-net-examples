using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the existing DOC/DOCX document.
        Document doc = new Document("Input.docx");

        // Find all shapes that are text boxes.
        List<Shape> textBoxShapes = new List<Shape>();
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each identified text box from the document.
        foreach (Shape shape in textBoxShapes)
            shape.Remove();

        // Prepare XPS save options (optional optimization).
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        xpsOptions.OptimizeOutput = true; // removes redundant canvases, concatenates runs, etc.

        // Save the modified document as XPS.
        doc.Save("Output.xps", xpsOptions);
    }
}
