using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class DeleteOleObjectExample
{
    static void Main()
    {
        // Load the DOCX document that contains OLE objects.
        Document doc = new Document("InputDocument.docx");

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove any shape that has an OLE object.
        foreach (Shape shape in shapes)
        {
            // The OleFormat property is non‑null only for OLE objects.
            if (shape.OleFormat != null)
            {
                // Remove the shape (and thus the embedded OLE object) from its parent.
                shape.Remove();
            }
        }

        // Save the modified document.
        doc.Save("OutputDocument.docx");
    }
}
