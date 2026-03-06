using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load the source document that contains an image.
        Document doc = new Document("Images.docx");

        // Retrieve the first shape that has an image.
        Shape imageShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];

        // Crop 30% from each side of the image.
        imageShape.ImageData.CropLeft = 0.3;
        imageShape.ImageData.CropRight = 0.3;
        imageShape.ImageData.CropTop = 0.3;
        imageShape.ImageData.CropBottom = 0.3;

        // Save the document with the cropped image.
        doc.Save("CroppedImage.docx");
    }
}
