using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

class BatchDocxToPdfWithWatermark
{
    static void Main()
    {
        // Base directory of the application.
        string baseDir = AppContext.BaseDirectory;

        // Folder that contains the source DOCX files.
        string inputFolder = Path.Combine(baseDir, "Input");

        // Folder where the resulting PDF files will be written.
        string outputFolder = Path.Combine(baseDir, "Output");

        // Path to the corporate logo that will be used as an image watermark.
        string logoPath = Path.Combine(baseDir, "CorporateLogo.png");

        // Ensure the input and output directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Ensure a placeholder logo image exists.
        if (!File.Exists(logoPath))
        {
            // A tiny 1x1 pixel PNG (transparent) encoded in base64.
            const string base64Png = 
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(logoPath, pngBytes);
        }

        // Ensure at least one DOCX file exists for processing.
        if (Directory.GetFiles(inputFolder, "*.docx").Length == 0)
        {
            string sampleDocx = Path.Combine(inputFolder, "Sample.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("This is a sample document generated for the batch conversion demo.");
            doc.Save(sampleDocx);
        }

        // Configure the watermark appearance (optional).
        var watermarkOptions = new ImageWatermarkOptions
        {
            // Scale the logo to 30 % of the page width (adjust as needed).
            Scale = 0.3,
            // Do not wash out the image; keep original colors.
            IsWashout = false
        };

        // Process each DOCX file in the input folder.
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            var document = new Document(docxFile);

            // Apply the corporate logo as an image watermark.
            document.Watermark.SetImage(logoPath, watermarkOptions);

            // Determine the output PDF file name.
            string pdfFile = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(docxFile) + ".pdf");

            // Save the document as PDF.
            document.Save(pdfFile);
        }

        Console.WriteLine("Batch conversion completed successfully.");
    }
}
