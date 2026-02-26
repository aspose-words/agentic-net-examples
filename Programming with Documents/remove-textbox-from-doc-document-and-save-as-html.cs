using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveHtml
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove any shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the modified document as HTML.
        doc.Save("OutputDocument.html", SaveFormat.Html);
    }
}
