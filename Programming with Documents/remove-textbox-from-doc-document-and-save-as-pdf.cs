using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find all shapes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove any shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as PDF.
        doc.Save("OutputDocument.pdf", SaveFormat.Pdf);
    }
}
