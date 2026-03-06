using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load an existing DOCX that contains an OLE object.
        Document doc = new Document("Input.docx");

        // Locate the first shape that holds an OLE object.
        Shape oleShape = doc.GetChild(NodeType.Shape, 0, true) as Shape;

        // If such a shape exists, remove it from the document.
        if (oleShape != null && oleShape.OleFormat != null)
        {
            oleShape.Remove(); // Deletes the OLE object together with its container shape.
        }

        // Save the document after the OLE object has been removed.
        doc.Save("Output.docx");
    }
}
