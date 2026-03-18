using System;
using System.IO;
using System.IO.Compression;
using Aspose.Words;
using Aspose.Words.Drawing;

class ConvertJpegToGrayscaleBmp
{
    static void Main()
    {
        // Paths (adjust as needed)
        string sourceDocPath = @"C:\Input\Images.docx";          // Document containing JPEG images
        string outputFolder = @"C:\Output\GrayscaleBmp\";       // Folder for BMP files
        string archivePath = @"C:\Output\GrayscaleImages.zip"; // Secure archive path

        // Ensure output directory exists
        Directory.CreateDirectory(outputFolder);

        // Load the source document
        Document doc = new Document(sourceDocPath);

        // Iterate through all shapes that contain images
        int imageIndex = 0;
        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
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
