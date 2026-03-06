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
        string pdfPath = @"C:\Input\sample.pdf";

        // Folder where extracted images will be saved.
        string outputFolder = @"C:\Output\ExtractedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document. No need to skip images, so we use default options.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        Document doc = new Document(pdfPath, loadOptions);

        // Get all Shape nodes in the document (including those inside headers/footers).
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain an image.
            if (shape.HasImage)
            {
                // Determine the appropriate file extension based on the image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string imageFileName = Path.Combine(outputFolder,
                    $"Image_{imageIndex}{extension}");

                // Save the image data to the file system.
                shape.ImageData.Save(imageFileName);

                imageIndex++;
            }
        }

        Console.WriteLine($"Extracted {imageIndex} image(s) to \"{outputFolder}\".");
    }
}
