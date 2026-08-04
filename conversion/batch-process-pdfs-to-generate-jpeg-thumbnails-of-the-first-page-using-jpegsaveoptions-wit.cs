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

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample PDF files to act as the batch source.
        for (int i = 1; i <= 3; i++)
        {
            Document sampleDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(sampleDoc);
            builder.Writeln($"Sample PDF document {i}");
            string pdfPath = Path.Combine(inputFolder, $"sample{i}.pdf");
            sampleDoc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Process each PDF in the input folder.
        foreach (string pdfFile in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfFile);

            // Configure image save options to render the first page as a low‑quality JPEG.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),

                // Low quality to increase compression (range 0‑100).
                JpegQuality = 10
            };

            // Determine the output JPEG file name.
            string jpegFile = Path.Combine(outputFolder,
                Path.GetFileNameWithoutExtension(pdfFile) + ".jpg");

            // Save the thumbnail.
            pdfDoc.Save(jpegFile, jpegOptions);

            // Validate that the thumbnail was created.
            if (!File.Exists(jpegFile) || new FileInfo(jpegFile).Length == 0)
                throw new InvalidOperationException($"Thumbnail was not created for '{pdfFile}'.");
        }

        // Optional: indicate completion.
        Console.WriteLine("Batch thumbnail generation completed successfully.");
    }
}
