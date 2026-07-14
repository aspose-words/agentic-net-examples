using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class BatchBmpToJpegConverter
{
    public static void Main()
    {
        // Define folders for input BMP images and output JPEG images.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputImages");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample BMP images if none exist.
        CreateSampleBmpImages(inputFolder);

        // Get all BMP files from the input folder.
        string[] bmpFiles = Directory.GetFiles(inputFolder, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found to convert.");

        // Batch convert each BMP to JPEG with 80% quality.
        foreach (string bmpPath in bmpFiles)
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the BMP image into the document.
            builder.InsertImage(bmpPath);

            // Configure JPEG save options with 80% quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80
            };

            // Determine output JPEG file path.
            string jpegFileName = Path.GetFileNameWithoutExtension(bmpPath) + ".jpg";
            string jpegPath = Path.Combine(outputFolder, jpegFileName);

            // Save the document page as a JPEG image.
            doc.Save(jpegPath, jpegOptions);

            // Log the conversion result.
            Console.WriteLine($"Converted '{Path.GetFileName(bmpPath)}' to '{jpegFileName}' with 80% quality.");
        }

        // Validate that at least one JPEG was produced.
        int jpegCount = Directory.GetFiles(outputFolder, "*.jpg").Length;
        if (jpegCount == 0)
            throw new InvalidOperationException("Conversion failed: no JPEG files were created.");
    }

    // Creates deterministic sample BMP images in the specified folder.
    private static void CreateSampleBmpImages(string folder)
    {
        // Define sample image specifications.
        var specs = new (int Width, int Height, Aspose.Drawing.Color Color, string Name)[]
        {
            (200, 200, Aspose.Drawing.Color.Red, "RedSquare.bmp"),
            (200, 200, Aspose.Drawing.Color.Green, "GreenSquare.bmp"),
            (200, 200, Aspose.Drawing.Color.Blue, "BlueSquare.bmp")
        };

        foreach (var spec in specs)
        {
            string filePath = Path.Combine(folder, spec.Name);
            if (File.Exists(filePath))
                continue; // Skip if already exists.

            // Create bitmap and draw a filled rectangle.
            using (Bitmap bitmap = new Bitmap(spec.Width, spec.Height))
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(spec.Color);
                bitmap.Save(filePath);
            }
        }
    }
}
