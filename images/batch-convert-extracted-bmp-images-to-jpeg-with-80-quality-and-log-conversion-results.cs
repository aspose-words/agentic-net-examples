using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample BMP images.
        for (int i = 1; i <= 3; i++)
        {
            string bmpPath = Path.Combine(inputDir, $"sample{i}.bmp");
            using (Bitmap bitmap = new Bitmap(200, 200))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                // Fill with a deterministic color.
                g.Clear(Color.FromArgb(50 * i, 80 * i, 120));
                // Save as BMP.
                bitmap.Save(bmpPath, ImageFormat.Bmp);
            }
        }

        // Batch convert BMP to JPEG with 80% quality.
        string[] bmpFiles = Directory.GetFiles(inputDir, "*.bmp");
        if (bmpFiles.Length == 0)
            throw new InvalidOperationException("No BMP files found for conversion.");

        foreach (string bmpFile in bmpFiles)
        {
            // Load image into a temporary document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertImage(bmpFile);

            // Configure JPEG save options.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 80
            };

            // Determine output path.
            string jpegFile = Path.Combine(outputDir,
                Path.GetFileNameWithoutExtension(bmpFile) + ".jpg");

            // Save the document page as a JPEG image.
            doc.Save(jpegFile, jpegOptions);

            // Log conversion result.
            FileInfo info = new FileInfo(jpegFile);
            Console.WriteLine($"Converted '{Path.GetFileName(bmpFile)}' to '{Path.GetFileName(jpegFile)}' – {info.Length} bytes.");
        }

        // Validate that at least one JPEG was produced.
        if (Directory.GetFiles(outputDir, "*.jpg").Length == 0)
            throw new InvalidOperationException("JPEG conversion failed; no output files were created.");
    }
}
