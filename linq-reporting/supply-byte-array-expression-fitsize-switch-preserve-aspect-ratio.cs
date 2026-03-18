using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertImageFromByteArray
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Load the image file into a byte array.
        // Replace the path with the actual location of your image.
        byte[] imageBytes = File.ReadAllBytes(@"C:\Images\Sample.jpg");

        // Insert the image as an inline shape. The InsertImage method returns the Shape object.
        Shape imageShape = builder.InsertImage(imageBytes);

        // Preserve the original aspect ratio by fitting the image data to the shape frame.
        // This scales the image so that it matches the shape's dimensions without distortion.
        imageShape.ImageData.FitImageToShape();

        // Save the document to a file.
        doc.Save(@"C:\Output\ImageFromByteArray.docx");
    }
}
