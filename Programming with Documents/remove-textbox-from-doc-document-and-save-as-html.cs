using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndConvert
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the collection indexing.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];
            // Remove the shape if it is a TextBox.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as HTML.
        doc.Save("OutputDocument.html", SaveFormat.Html);
    }
}
