using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample PDF document.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Sample PDF content for conversion.");
        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        sampleDoc.Save(pdfPath, SaveFormat.Pdf);

        // Load the PDF document to be converted.
        Document pdfDoc = new Document(pdfPath);

        // Configure image save options for TIFF output.
        // ImageContrast must be in the range 0..1; using 0.8 for high contrast.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Lzw,
            ImageContrast = 0.8f // High contrast within valid range.
        };

        // Save the PDF as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "converted.tiff");
        pdfDoc.Save(tiffPath, tiffOptions);

        // Validate that the TIFF file was created and is not empty.
        if (!File.Exists(tiffPath) || new FileInfo(tiffPath).Length == 0)
        {
            throw new InvalidOperationException("TIFF conversion failed: output file is missing or empty.");
        }

        // Indicate successful completion.
        Console.WriteLine("PDF successfully converted to TIFF with LZW compression and custom contrast.");
    }
}
