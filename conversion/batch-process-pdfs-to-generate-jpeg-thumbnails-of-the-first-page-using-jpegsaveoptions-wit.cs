using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input PDFs and output JPEG thumbnails.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Thumbnails");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample PDF files to work with.
        for (int i = 1; i <= 3; i++)
        {
            string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"This is sample PDF number {i}.");
            // Save as PDF.
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Process each PDF in the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf");
        foreach (string pdfFile in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure image save options for JPEG with low quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),
                // Low quality to increase compression (range 0‑100).
                JpegQuality = 10,
                // Optional: set a reasonable resolution.
                Resolution = 150
            };

            // Determine the output JPEG file path.
            string outputFileName = Path.GetFileNameWithoutExtension(pdfFile) + ".jpg";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the first page as a JPEG thumbnail.
            pdfDoc.Save(outputPath, jpegOptions);

            // Validate that the thumbnail was created.
            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            {
                throw new InvalidOperationException($"Thumbnail was not created for '{pdfFile}'.");
            }
        }

        // Optionally, indicate completion (no interactive prompts).
        Console.WriteLine($"Processed {pdfFiles.Length} PDF(s). Thumbnails are saved in '{outputFolder}'.");
    }
}
