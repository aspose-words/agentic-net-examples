using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("Input.docx");

        // Find all shapes that are text boxes.
        List<Shape> textBoxes = new List<Shape>();
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.TextBox)
                textBoxes.Add(shape);
        }

        // Remove each text box from the document.
        foreach (Shape tb in textBoxes)
            tb.Remove();

        // Save the resulting document as plain text.
        TxtSaveOptions txtOptions = new TxtSaveOptions();
        doc.Save("Output.txt", txtOptions);
    }
}
