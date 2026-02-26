using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class RemoveTextBoxExample
{
    static void Main()
    {
        // Load the source document (DOC or DOCX).
        Document doc = new Document("Input.docx");

        // Retrieve all shape nodes (including text boxes) in the document.
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate backwards so that removal does not affect the loop indexing.
        for (int i = shapes.Count - 1; i >= 0; i--)
        {
            Shape shape = (Shape)shapes[i];

            // Identify text box shapes and remove them.
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Save the result as a DOC file.
        doc.Save("Output.doc");
    }
}
