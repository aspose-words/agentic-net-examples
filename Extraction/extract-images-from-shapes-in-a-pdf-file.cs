using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;

class ExtractImagesFromPdf
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\sample.pdf";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Docs\ExtractedImages\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the PDF document. Set SkipPdfImages to false to load images.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            SkipPdfImages = false
        };
        Document doc = new Document(pdfPath, loadOptions);

        // Retrieve all shape nodes (which may contain images).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string fileName = Path.Combine(outputDir, $"Image_{imageIndex}{extension}");

                // Save the image to the file system.
                shape.ImageData.Save(fileName);
                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
