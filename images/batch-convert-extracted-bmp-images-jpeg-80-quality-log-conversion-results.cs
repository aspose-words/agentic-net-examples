using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class BmpToJpegBatchConverter
{
    // Adjust these paths as needed.
    private const string InputFolder = @"C:\Images\Bmp";
    private const string OutputFolder = @"C:\Images\Jpeg";

    static void Main()
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(OutputFolder);

        // Get all BMP files in the input folder.
        string[] bmpFiles = Directory.GetFiles(InputFolder, "*.bmp", SearchOption.TopDirectoryOnly);

        foreach (string bmpPath in bmpFiles)
        {
            try
            {
                // Create a new empty Word document.
                Document doc = new Document();

                // Insert the BMP image into the document.
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.InsertImage(bmpPath);

                // Prepare save options for JPEG with 80% quality.
                ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
                {
                    JpegQuality = 80 // 80% quality as required.
                };

                // Determine the output JPEG file name.
                string fileNameWithoutExt = Path.GetFileNameWithoutExtension(bmpPath);
                string jpegPath = Path.Combine(OutputFolder, fileNameWithoutExt + ".jpg");

                // Save the document page (which contains only the image) as JPEG.
                doc.Save(jpegPath, jpegOptions);

                // Log successful conversion.
                Console.WriteLine($"Converted: {Path.GetFileName(bmpPath)} -> {Path.GetFileName(jpegPath)} (Quality: 80)");
            }
            catch (Exception ex)
            {
                // Log any errors that occur during conversion.
                Console.WriteLine($"Failed to convert {Path.GetFileName(bmpPath)}: {ex.Message}");
            }
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
