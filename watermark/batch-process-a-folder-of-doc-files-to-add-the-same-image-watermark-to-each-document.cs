using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Base directories for input documents, output documents and the watermark image.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputDocs");
        string imagePath = Path.Combine(baseDir, "watermark.png");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a simple PNG image to be used as the watermark.
        // -----------------------------------------------------------------
        // This is a 1x1 pixel transparent PNG.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(imagePath, pngBytes);

        // -----------------------------------------------------------------
        // 2. Create a few sample DOCX files to demonstrate batch processing.
        // -----------------------------------------------------------------
        for (int i = 1; i <= 3; i++)
        {
            string docPath = Path.Combine(inputDir, $"Sample{i}.docx");
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln($"This is sample document #{i}.");
            doc.Save(docPath);
        }

        // -----------------------------------------------------------------
        // 3. Prepare watermark options (optional: scale and washout).
        // -----------------------------------------------------------------
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            Scale = 5,          // Scale factor (example value).
            IsWashout = false   // Do not apply washout effect.
        };

        // -----------------------------------------------------------------
        // 4. Process each DOC/DOCX file in the input folder.
        // -----------------------------------------------------------------
        string[] docFiles = Directory.GetFiles(inputDir, "*.docx");
        foreach (string inputFile in docFiles)
        {
            // Load the document.
            Document document = new Document(inputFile);

            // Apply the image watermark using the file path and options.
            document.Watermark.SetImage(imagePath, watermarkOptions);

            // Determine output file name.
            string fileName = Path.GetFileNameWithoutExtension(inputFile);
            string outputFile = Path.Combine(outputDir, $"{fileName}_Watermarked.docx");

            // Save the watermarked document.
            document.Save(outputFile);
        }

        // -----------------------------------------------------------------
        // 5. Simple validation: ensure that output files were created.
        // -----------------------------------------------------------------
        foreach (string outFile in Directory.GetFiles(outputDir, "*_Watermarked.docx"))
        {
            Console.WriteLine($"Created watermarked file: {outFile}");
        }
    }
}
