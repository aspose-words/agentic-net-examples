using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffToGrayscaleJpegConverter
{
    // Convert all TIFF files in a source folder to grayscale JPEG files in a destination folder.
    static void Main()
    {
        // Folder that contains the source TIFF images.
        string sourceFolder = @"C:\Images\Tiff";
        // Folder where the resulting JPEG images will be saved.
        string destinationFolder = @"C:\Images\Jpeg";

        // Ensure the destination folder exists.
        Directory.CreateDirectory(destinationFolder);

        // Get all *.tif and *.tiff files from the source folder.
        string[] tiffFiles = Directory.GetFiles(sourceFolder, "*.tif");
        string[] tiffFilesAlt = Directory.GetFiles(sourceFolder, "*.tiff");
        string[] allTiffFiles = new string[tiffFiles.Length + tiffFilesAlt.Length];
        tiffFiles.CopyTo(allTiffFiles, 0);
        tiffFilesAlt.CopyTo(allTiffFiles, tiffFiles.Length);

        foreach (string tiffPath in allTiffFiles)
        {
            // Load the TIFF image into a temporary Word document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the TIFF image. Aspose.Words will embed the image as a shape.
            builder.InsertImage(tiffPath);

            // Prepare save options for JPEG output.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render the image in grayscale to reduce size.
                ImageColorMode = ImageColorMode.Grayscale,
                // Adjust JPEG quality for low‑bandwidth scenarios (0‑100).
                JpegQuality = 50
            };

            // Build the output file name (same base name, .jpg extension).
            string outputFileName = Path.GetFileNameWithoutExtension(tiffPath) + ".jpg";
            string outputPath = Path.Combine(destinationFolder, outputFileName);

            // Save the document page (which contains only the inserted image) as a JPEG.
            doc.Save(outputPath, saveOptions);
        }
    }
}
