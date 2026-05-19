using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string inputDir = "InputImages";
        string outputDir = "OutputImages";
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample BMP images
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            using (Bitmap bitmap = new Bitmap(100, 100))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with a distinct color for each image
                int r = (i * 70) % 256;
                int gVal = (i * 130) % 256;
                int b = (i * 200) % 256;
                g.Clear(Color.FromArgb(r, gVal, b));
                bitmap.Save(bmpPath);
            }
        }

        int convertedCount = 0;

        // Batch convert BMP to JPEG with 80% quality
        foreach (string bmpFile in Directory.GetFiles(inputDir, "*.bmp"))
        {
            // Load BMP into a temporary document
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(bmpFile);

            // Configure JPEG save options
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80
            };

            // Determine output path
            string outputFile = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(bmpFile) + ".jpg");

            // Save the document page (containing the image) as JPEG
            doc.Save(outputFile, jpegOptions);

            // Validate output
            if (!File.Exists(outputFile))
                throw new InvalidOperationException($"Conversion failed for {bmpFile}");

            Console.WriteLine($"Converted '{Path.GetFileName(bmpFile)}' to '{Path.GetFileName(outputFile)}'.");
            convertedCount++;
        }

        // Ensure at least one image was processed
        if (convertedCount == 0)
            throw new InvalidOperationException("No BMP images were found for conversion.");
    }
}
