using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class RemoveTextBoxAndSaveAsPs
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.doc");

        // Find all Shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the loop index.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            // Remove the shape if it is a text box.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Prepare PostScript save options.
        PsSaveOptions psOptions = new PsSaveOptions
        {
            SaveFormat = SaveFormat.Ps // Explicitly set the format to PostScript.
        };

        // Save the modified document as a .ps file.
        doc.Save("OutputDocument.ps", psOptions);
    }
}
