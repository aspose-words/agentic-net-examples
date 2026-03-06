using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Input WORDML document, image to use as watermark, and output file paths.
        string inputPath = "Input.docx";
        string imagePath = "Logo.jpg";
        string outputPath = "Output.docx";

        // Load the existing document.
        Document doc = new Document(inputPath);

        // Configure image watermark appearance.
        ImageWatermarkOptions options = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor of the watermark image.
            IsWashout = false   // Preserve original image colors.
        };

        // Add the image watermark to the document.
        doc.Watermark.SetImage(imagePath, options);

        // Save the document with the watermark applied.
        doc.Save(outputPath);
    }
}
