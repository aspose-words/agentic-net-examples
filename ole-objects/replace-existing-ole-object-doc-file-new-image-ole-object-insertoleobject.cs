using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class ReplaceOleObject
{
    static void Main()
    {
        // Paths to the source document, the new image file and the output document.
        string sourceDocPath = @"C:\Docs\Original.docx";
        string newImagePath = @"C:\Images\Replacement.png";
        string outputDocPath = @"C:\Docs\Modified.docx";

        // Load the existing Word document.
        Document doc = new Document(sourceDocPath);

        // Locate the first shape that contains an OLE object.
        Shape oleShape = null;
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.OleObject)
            {
                oleShape = shape;
                break;
            }
        }

        if (oleShape == null)
        {
            Console.WriteLine("No OLE object found in the document.");
            return;
        }

        // Remove the existing OLE shape.
        oleShape.Remove();

        // Move the builder to the position where the old shape was.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveTo(oleShape);

        // Insert the new image as an OLE object.
        // Use the "Package" progId to embed arbitrary file data.
        // The image will be displayed as its content (asIcon = false) and no custom presentation image is supplied.
        using (FileStream imageStream = File.OpenRead(newImagePath))
        {
            builder.InsertOleObject(imageStream, "Package", false, null);
        }

        // Save the modified document.
        doc.Save(outputDocPath);
    }
}
