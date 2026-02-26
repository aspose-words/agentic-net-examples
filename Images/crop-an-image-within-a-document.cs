using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class CropImageExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an image into the document. The InsertImage method returns the Shape that contains the image.
        // Replace "ImageDir + \"Sample.jpg\"" with the actual path to your image file.
        Shape imageShape = builder.InsertImage(@"C:\Images\Sample.jpg");

        // Access the ImageData of the shape to modify cropping.
        ImageData imgData = imageShape.ImageData;

        // Crop 20% from the left, 10% from the top, 15% from the right, and 5% from the bottom.
        imgData.CropLeft = 0.20;   // 20% removal from the left side.
        imgData.CropTop = 0.10;    // 10% removal from the top side.
        imgData.CropRight = 0.15;  // 15% removal from the right side.
        imgData.CropBottom = 0.05; // 5% removal from the bottom side.

        // Save the modified document to a file.
        doc.Save(@"C:\Output\CroppedImage.docx");
    }
}
