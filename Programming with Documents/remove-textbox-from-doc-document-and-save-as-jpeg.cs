using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

class RemoveTextBoxAndConvertToJpeg
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("InputDocument.doc");

        // Find all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

        // Remove every shape that is a TextBox.
        foreach (Shape shape in shapes)
        {
            if (shape.ShapeType == ShapeType.TextBox)
                shape.Remove();
        }

        // Prepare image save options for JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Save the modified document as a JPEG image (first page only).
        doc.Save("OutputImage.jpg", jpegOptions);
    }
}
