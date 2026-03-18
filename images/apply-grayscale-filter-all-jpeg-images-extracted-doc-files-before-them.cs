using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class GrayscaleJpegProcessor
{
    static void Main()
    {
        // Path to the source DOC/DOCX file.
        string inputFile = @"C:\Docs\InputDocument.docx";

        // Path where the processed document will be saved.
        string outputFile = @"C:\Docs\OutputDocument.docx";

        // Load the document (create/load rule).
        Document doc = new Document(inputFile);

        // Iterate through all shapes in the document.
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
        {
            // Process only shapes that contain an image.
            if (!shape.HasImage)
                continue;

            ImageData imageData = shape.ImageData;

            // Apply the grayscale filter only to JPEG images.
            if (imageData.ImageType == ImageType.Jpeg)
            {
                // Set the GrayScale property to true (feature rule).
                imageData.GrayScale = true;
            }
        }

        // Save the modified document (save rule).
        doc.Save(outputFile);
    }
}
