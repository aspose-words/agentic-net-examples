using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class CropImageExample
{
    static void Main()
    {
        // Load an existing document that contains an image.
        // Replace {InputPath} with the path to your source .docx file.
        Document doc = new Document("{InputPath}");

        // Find the first shape that contains an image.
        // Shapes are stored as nodes of type Shape.
        Shape imageShape = (Shape)doc.GetChildNodes(NodeType.Shape, true)[0];

        // Crop the image by setting the fractional values (0.0 – 1.0).
        // Positive values remove that fraction from the corresponding side.
        imageShape.ImageData.CropLeft = 0.2;   // Remove 20% from the left.
        imageShape.ImageData.CropRight = 0.2;  // Remove 20% from the right.
        imageShape.ImageData.CropTop = 0.1;    // Remove 10% from the top.
        imageShape.ImageData.CropBottom = 0.1; // Remove 10% from the bottom.

        // Save the modified document.
        // Replace {OutputPath} with the desired output file location.
        doc.Save("{OutputPath}");
    }
}
