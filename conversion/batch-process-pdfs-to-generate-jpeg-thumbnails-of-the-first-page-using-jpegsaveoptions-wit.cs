using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders for input PDFs and output thumbnails.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputPdfs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Thumbnails");
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample PDF files.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document #{i}");
            builder.Writeln("This document is generated programmatically for thumbnail extraction.");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Process each PDF and generate a low‑quality JPEG thumbnail of the first page.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure image save options for JPEG with low quality.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                JpegQuality = 10,               // Low quality for higher compression.
                PageSet = new PageSet(0)        // Render only the first page.
            };

            // Determine output thumbnail path.
            string thumbnailPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(pdfFile) + ".jpg");

            // Save the thumbnail.
            pdfDoc.Save(thumbnailPath, jpegOptions);

            // Validate that the thumbnail was created.
            if (!File.Exists(thumbnailPath) || new FileInfo(thumbnailPath).Length == 0)
                throw new InvalidOperationException($"Thumbnail was not created for '{pdfFile}'.");
        }

        // Indicate successful completion.
        Console.WriteLine("Thumbnails generated successfully.");
    }
}
