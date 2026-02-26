using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the existing DOC/DOCX document.
        Document doc = new Document("input.docx");

        // Collect all shapes that are text boxes.
        List<Shape> textBoxShapes = new List<Shape>();
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxShapes.Add(shape);
        }

        // Remove each text box from the document.
        foreach (Shape tb in textBoxShapes)
            tb.Remove();

        // Save the modified document as EPUB. The format is inferred from the file extension.
        doc.Save("output.epub");
    }
}
