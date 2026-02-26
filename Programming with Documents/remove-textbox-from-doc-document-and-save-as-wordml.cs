using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxAndSaveAsWordML
{
    static void Main()
    {
        // Load the existing DOC document.
        Document doc = new Document("InputDocument.docx");

        // Find all shape nodes in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the collection in reverse order to safely remove nodes.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Check if the shape is a TextBox.
            if (shape.ShapeType == ShapeType.TextBox)
            {
                // Remove the textbox from its parent node.
                shape.Remove();
            }
        }

        // Save the modified document as WORDML (XML format).
        doc.Save("OutputDocument.xml", SaveFormat.WordML);
    }
}
