using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Drawing;

class ExtractImagesFromPdf
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Directory where extracted images will be saved.
        string outputDir = @"C:\Output\Images";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the PDF document. Ensure images are not skipped.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            SkipPdfImages = false
        };
        Document doc = new Document(pdfPath, loadOptions);

        // Retrieve all shape nodes from the document (including image shapes).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension for the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the full file name for the extracted image.
                string imageFilePath = Path.Combine(outputDir, $"Image_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imageFilePath);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputDir}\".");
    }
}
