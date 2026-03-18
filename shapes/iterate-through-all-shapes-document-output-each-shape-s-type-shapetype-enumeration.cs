using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Retrieve all Shape nodes in the document (including those inside groups).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        // Iterate through each Shape and output its ShapeType enumeration value.
        foreach (Shape shape in shapeNodes)
        {
            Console.WriteLine(shape.ShapeType);
        }

        // Save the document if any modifications were made (optional).
        doc.Save("Output.docx");
    }
}
