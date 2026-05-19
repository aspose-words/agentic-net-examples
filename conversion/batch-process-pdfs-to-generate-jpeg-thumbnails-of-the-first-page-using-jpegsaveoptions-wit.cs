using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string inputFolder = "InputPdfs";
        string outputFolder = "OutputThumbnails";
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample PDF files.
        CreateSamplePdf(Path.Combine(inputFolder, "Sample1.pdf"));
        CreateSamplePdf(Path.Combine(inputFolder, "Sample2.pdf"));

        // Process each PDF in the input folder.
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Configure image save options for a low‑quality JPEG thumbnail of the first page.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the first page (zero‑based index).
                PageSet = new PageSet(0),
                // Low quality (strong compression).
                JpegQuality = 10
            };

            // Build the output file name.
            string outputFileName = Path.GetFileNameWithoutExtension(pdfPath) + "_thumb.jpg";
            string outputPath = Path.Combine(outputFolder, outputFileName);

            // Save the thumbnail.
            pdfDoc.Save(outputPath, jpegOptions);

            // Verify that the thumbnail was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException($"Thumbnail was not created: {outputPath}");
        }
    }

    // Helper method to create a simple two‑page PDF.
    private static void CreateSamplePdf(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("First page of the sample PDF.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the sample PDF.");

        doc.Save(filePath, SaveFormat.Pdf);
    }
}
