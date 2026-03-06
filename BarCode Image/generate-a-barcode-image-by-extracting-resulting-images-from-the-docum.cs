using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Drawing;

namespace BarcodeImageExtraction
{
    // Simple custom barcode generator that returns a minimal PNG image.
    // This implementation avoids System.Drawing dependencies, making the code compile on all .NET platforms.
    public class CustomBarcodeGenerator : IBarcodeGenerator
    {
        // Generates an image for DISPLAYBARCODE fields.
        public Stream GetBarcodeImage(BarcodeParameters parameters)
        {
            // A 1x1 pixel transparent PNG (base64 encoded).
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            return new MemoryStream(pngBytes);
        }

        // Generates an image for old‑fashioned BARCODE fields.
        public Stream GetOldBarcodeImage(BarcodeParameters parameters)
        {
            // Reuse the same implementation for simplicity.
            return GetBarcodeImage(parameters);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the input DOCX that contains DISPLAYBARCODE fields.
            string inputPath = @"ArtifactsDir\InputWithBarcodes.docx";
            // Directory where extracted images will be saved.
            string outputDir = @"ArtifactsDir\ExtractedImages";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDir);

            // Load the document.
            Document doc = new Document(inputPath);

            // Assign the custom barcode generator so that field updates produce images.
            doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();

            // Update all fields in the document – this will generate barcode images.
            doc.UpdateFields();

            // Iterate through all Shape nodes that contain images.
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;

            foreach (Shape shape in shapes)
            {
                if (shape.IsImage)
                {
                    // Get the raw image bytes.
                    byte[] imageBytes = shape.ImageData.ImageBytes;

                    // Determine the image format from the shape's image type.
                    ImageType imageType = shape.ImageData.ImageType;
                    string extension = imageType switch
                    {
                        ImageType.Jpeg => "jpg",
                        ImageType.Png => "png",
                        ImageType.Gif => "gif",
                        ImageType.Bmp => "bmp",
                        ImageType.Emf => "emf",
                        ImageType.Wmf => "wmf",
                        _ => "bin"
                    };

                    // Save the image to a file.
                    string imagePath = Path.Combine(outputDir, $"BarcodeImage_{imageIndex}.{extension}");
                    File.WriteAllBytes(imagePath, imageBytes);
                    imageIndex++;
                }
            }

            // Optionally, save the updated document (with generated barcode images) to a new file.
            string updatedDocPath = @"ArtifactsDir\DocumentWithGeneratedBarcodes.docx";
            doc.Save(updatedDocPath);
        }
    }
}
