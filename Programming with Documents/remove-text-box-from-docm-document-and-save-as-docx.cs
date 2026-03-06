using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the macro‑enabled document (DOCM)
        Document doc = new Document("Input.docm");

        // Remove every shape that is a text box
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the cleaned document as a macro‑free DOCX
        doc.Save("Output.docx", SaveFormat.Docx);
    }
}
