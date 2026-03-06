using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsDotm
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Iterate through all Shape nodes in the document.
        // Shapes with ShapeType.TextBox represent text boxes.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the text box from its parent node.
                shape.Remove();
            }
        }

        // Save the modified document as a DOTM (macro‑enabled template) file.
        doc.Save("OutputDocument.dotm", SaveFormat.Dotm);
    }
}
