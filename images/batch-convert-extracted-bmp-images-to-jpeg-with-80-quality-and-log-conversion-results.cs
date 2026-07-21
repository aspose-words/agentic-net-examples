using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchBmpToJpegConverter
{
    public static void Main()
    {
        // Prepare input and output folders.
        string inputFolder = "InputImages";
        string outputFolder = "OutputImages";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images if none exist.
        CreateSampleBmpImages(inputFolder, 3);

        // Convert each BMP image to JPEG with 80% quality.
        int convertedCount = 0;
        foreach (string bmpPath in Directory.GetFiles(inputFolder, "*.bmp"))
        {
            // Load the BMP into a new blank document and insert the image.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(bmpPath);

            // Configure JPEG save options with the required quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80
            };

            // Determine output file name.
            string jpegPath = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(bmpPath) + ".jpg");

            // Save the document page (containing only the image) as JPEG.
            doc.Save(jpegPath, jpegOptions);

            Console.WriteLine($"Converted '{bmpPath}' to '{jpegPath}' with 80% quality.");
            convertedCount++;
        }

        // Validate that at least one image was processed.
        if (convertedCount == 0)
            throw new Exception("No BMP images were found for conversion.");
    }

    // Generates a specified number of deterministic BMP files in the given folder.
    private static void CreateSampleBmpImages(string folder, int count)
    {
        for (int i = 1; i <= count; i++)
        {
            string filePath = Path.Combine(folder, $"sample{i}.bmp");
            if (File.Exists(filePath))
                continue; // Skip if already created.

            // Create a 100x100 bitmap with a simple colored background.
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill with a distinct color for each image.
                int red = (i * 50) % 256;
                int green = (i * 80) % 256;
                int blue = (i * 110) % 256;
                graphics.Clear(Color.FromArgb(red, green, blue));

                // Save as BMP.
                bitmap.Save(filePath);
            }
        }
    }
}
