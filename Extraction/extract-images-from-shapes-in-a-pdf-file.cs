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

        // Load the PDF document. PdfLoadOptions allows us to control loading behavior.
        var loadOptions = new PdfLoadOptions();
        Document doc = new Document(pdfPath, loadOptions);

        // Iterate through all shapes in the document.
        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
        int imageIndex = 0;

        foreach (Shape shape in shapes)
        {
            // Check that the shape actually contains an image.
            if (shape.HasImage && shape.ImageData.HasImage)
            {
                // Retrieve the raw image bytes.
                byte[] imageBytes = shape.ImageData.ImageBytes;

                // Determine a suitable file extension based on the image type.
                string extension = GetExtension(shape.ImageData.ImageType);

                // Build the output file name.
                string fileName = $"Image_{imageIndex}{extension}";
                string outputPath = Path.Combine(outputFolder, fileName);

                // Write the image to disk.
                File.WriteAllBytes(outputPath, imageBytes);

                imageIndex++;
            }
        }
    }

    // Maps Aspose.Words.ImageType values to common file extensions.
    private static string GetExtension(ImageType imageType)
    {
        switch (imageType)
        {
            case ImageType.Jpeg:
                return ".jpg";
            case ImageType.Png:
                return ".png";
            case ImageType.Bmp:
                return ".bmp";
            case ImageType.Gif:
                return ".gif";
            case ImageType.WebP:
                return ".webp";
            case ImageType.Emf:
                return ".emf";
            case ImageType.Wmf:
                return ".wmf";
            case ImageType.Pict:
                return ".pict";
            case ImageType.Eps:
                return ".eps";
            default:
                return ".bin"; // Fallback for unknown or unsupported types.
        }
    }
}
