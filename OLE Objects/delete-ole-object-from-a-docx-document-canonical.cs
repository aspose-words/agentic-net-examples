using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class DeleteOleObject
{
    static void Main()
    {
        // Path to the source DOCX file that contains OLE objects.
        string inputPath = "input.docx";

        // Path where the modified document will be saved.
        string outputPath = "output.docx";

        // Load the document (uses the Document(string) constructor rule).
        Document doc = new Document(inputPath);

        // Retrieve all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through the shapes and remove those that are OLE objects.
        foreach (Shape shape in shapeNodes)
        {
            if (shape.ShapeType == ShapeType.OleObject)
            {
                // Remove the OLE object shape from its parent node.
                shape.Remove();
            }
        }

        // Save the updated document (uses the Document.Save(string) rule).
        doc.Save(outputPath);
    }
}
