using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsWordML
{
    static void Main()
    {
        // Load an existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Retrieve all shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection and remove every shape that is a text box.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove(); // Removes the shape from its parent node.
        }

        // Save the modified document in WordML (XML) format.
        doc.Save("OutputDocument.xml", SaveFormat.WordML);
    }
}
