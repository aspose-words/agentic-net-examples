using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class JpegToWebpConverter
{
    static void Main()
    {
        // Folder containing the source JPEG images.
        string inputFolder = @"C:\Images\Jpeg";
        // Folder where the converted WebP images will be saved.
        string outputFolder = @"C:\Images\Webp";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Process each JPEG file in the input folder.
        foreach (string jpegPath in Directory.GetFiles(inputFolder, "*.jpg"))
        {
            // Create a new empty document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the JPEG image into the document.
            builder.InsertImage(jpegPath);

            // Configure image save options for WebP format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.WebP)
            {
                // Use high‑quality rendering algorithms for the best visual result.
                UseHighQualityRendering = true,
                // Optional: set the scale to 1 (default) to keep original dimensions.
                Scale = 1.0f
            };

            // Determine the output file name with .webp extension.
            string outputPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(jpegPath) + ".webp");

            // Save the document page (containing the image) as a WebP image.
            doc.Save(outputPath, saveOptions);
        }

        Console.WriteLine("Conversion completed.");
    }
}
