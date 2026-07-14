using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Define folders for input DOCX files, output PDFs and the logo image.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string logoPath = Path.Combine(baseDir, "logo.png");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create a simple logo image (a red 1x1 PNG) if it does not exist.
        // The image is stored as a Base64 string to avoid System.Drawing usage.
        // -----------------------------------------------------------------
        if (!File.Exists(logoPath))
        {
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8z8BQDwAF/AL+XcXK5wAAAABJRU5ErkJggg==";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(logoPath, pngBytes);
        }

        // -----------------------------------------------------------------
        // Create a few sample DOCX files if the input folder is empty.
        // -----------------------------------------------------------------
        if (Directory.GetFiles(inputDir, "*.docx").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                string samplePath = Path.Combine(inputDir, $"Sample{i}.docx");
                Document sampleDoc = new Document();
                var builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document #{i}.");
                sampleDoc.Save(samplePath);
            }
        }

        // -----------------------------------------------------------------
        // Process each DOCX: add image watermark and convert to PDF.
        // -----------------------------------------------------------------
        foreach (string docxFile in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the source document.
            Document doc = new Document(docxFile);

            // Configure image watermark options (optional).
            ImageWatermarkOptions imgOptions = new ImageWatermarkOptions
            {
                Scale = 0.5,          // 50% of the page width.
                IsWashout = false    // Keep original colors.
            };

            // Apply the image watermark using the logo file.
            doc.Watermark.SetImage(logoPath, imgOptions);

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxFile) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        Console.WriteLine("Batch conversion completed. PDFs are located in: " + outputDir);
    }
}
