using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffToJpegBatchConverter
{
    static void Main()
    {
        // Directory containing the source TIFF images.
        string inputDirectory = @"C:\InputTiffs";

        // Directory where the converted JPEG images will be saved.
        string outputDirectory = @"C:\OutputJpegs";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDirectory);

        // Iterate over all TIFF files in the input directory.
        foreach (string tiffFilePath in Directory.GetFiles(inputDirectory, "*.tif*"))
        {
            // Create a new empty Word document.
            Document doc = new Document();

            // Insert the TIFF image into the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(tiffFilePath);

            // Configure image save options for JPEG with 90% quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 90
            };

            // Build the output JPEG file path.
            string jpegFileName = Path.GetFileNameWithoutExtension(tiffFilePath) + ".jpg";
            string jpegFilePath = Path.Combine(outputDirectory, jpegFileName);

            // Save the document as a JPEG image using the configured options.
            doc.Save(jpegFilePath, jpegOptions);
        }
    }
}
