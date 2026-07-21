using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Directories for input DOCX files and output PDF files.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");

        // Path for the corporate logo image that will be used as a watermark.
        string logoPath = Path.Combine(baseDir, "logo.png");

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create a deterministic small PNG image (1x1 pixel, transparent) if it does not exist.
        if (!File.Exists(logoPath))
        {
            // Base64 representation of a 1x1 transparent PNG.
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(logoPath, pngBytes);
        }

        // Create sample DOCX files if the input folder is empty.
        string[] existingDocs = Directory.GetFiles(inputDir, "*.docx");
        if (existingDocs.Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample document #{i}.");
                string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
                sampleDoc.Save(docPath);
            }
        }

        // Process each DOCX file: add image watermark and convert to PDF.
        foreach (string docFilePath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docFilePath);

            // Configure image watermark options.
            ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
            {
                Scale = 5,          // Scale factor for the watermark image.
                IsWashout = false   // Disable washout effect to keep the logo vivid.
            };

            // Apply the image watermark using the logo file.
            doc.Watermark.SetImage(logoPath, watermarkOptions);

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputDir, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfFilePath, SaveFormat.Pdf);
        }

        // Optional: verify that PDFs were created (no console output required).
        // The program ends here.
    }
}
