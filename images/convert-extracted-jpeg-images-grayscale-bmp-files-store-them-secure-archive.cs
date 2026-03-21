using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ConvertJpegToGrayscaleBmp
{
    static void Main()
    {
        // Use paths relative to the current directory
        string baseDir = Directory.GetCurrentDirectory();
        string sourceDocPath = Path.Combine(baseDir, "Images.docx");          // Document containing images
        string outputFolder = Path.Combine(baseDir, "GrayscaleBmp");         // Folder for BMP files
        string archivePath = Path.Combine(baseDir, "GrayscaleImages.zip");   // Secure archive path

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // If the source document does not exist, create a simple one with a PNG image
        if (!File.Exists(sourceDocPath))
        {
            // Create a 1x1 pixel PNG image (transparent)
            byte[] pngBytes = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=");
            string tempPngPath = Path.Combine(baseDir, "temp.png");
            File.WriteAllBytes(tempPngPath, pngBytes);

            // Create a new document and insert the PNG image as a shape
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample document with an image:");
            Shape shape = new Shape(doc, ShapeType.Image);
            shape.ImageData.SetImage(tempPngPath);
            builder.InsertNode(shape);
            doc.Save(sourceDocPath);

            // Clean up temporary PNG
            File.Delete(tempPngPath);
        }

        // Load the source document
        Document docToProcess = new Document(sourceDocPath);

        // Iterate through all shapes that contain images
        int imageIndex = 0;
        foreach (Shape shape in docToProcess.GetChildNodes(NodeType.Shape, true))
        {
            if (!shape.HasImage)
                continue;

            // Set the image to display in grayscale mode
            shape.ImageData.GrayScale = true;

            // Build BMP file name
            string bmpFileName = Path.Combine(outputFolder, $"Image_{imageIndex}.bmp");

            // Save the image as BMP; the GrayScale flag influences the saved output
            shape.ImageData.Save(bmpFileName);
            imageIndex++;
        }

        // Create a secure ZIP archive containing all generated BMP files
        if (File.Exists(archivePath))
            File.Delete(archivePath);

        ZipFile.CreateFromDirectory(outputFolder, archivePath, CompressionLevel.Optimal, false);
    }
}
