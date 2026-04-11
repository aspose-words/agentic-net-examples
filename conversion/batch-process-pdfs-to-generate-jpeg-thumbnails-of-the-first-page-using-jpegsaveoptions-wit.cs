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

        // Create sample PDF files if the input folder is empty.
        if (Directory.GetFiles(inputFolder, "*.pdf").Length == 0)
        {
            for (int i = 1; i <= 3; i++)
            {
                Document sampleDoc = new Document();
                DocumentBuilder builder = new DocumentBuilder(sampleDoc);
                builder.Writeln($"This is sample PDF number {i}.");
                // Save as PDF.
                string pdfPath = Path.Combine(inputFolder, $"Sample{i}.pdf");
                sampleDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }

        // Process each PDF in the input folder.
        foreach (string pdfFilePath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFilePath);

            // Configure image save options for JPEG with low quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 10,               // Low quality for higher compression.
                PageSet = new PageSet(0)        // First page only.
            };

            // Determine output JPEG file path.
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(pdfFilePath);
            string jpegPath = Path.Combine(outputFolder, $"{fileNameWithoutExt}_thumb.jpg");

            // Save the first page as a JPEG thumbnail.
            pdfDoc.Save(jpegPath, jpegOptions);

            // Validate that the JPEG file was created and has content.
            if (!File.Exists(jpegPath) || new FileInfo(jpegPath).Length == 0)
            {
                throw new InvalidOperationException($"Failed to create thumbnail for '{pdfFilePath}'.");
            }
        }

        // Optional: indicate completion (no interactive prompts).
        Console.WriteLine("Batch thumbnail generation completed.");
    }
}
