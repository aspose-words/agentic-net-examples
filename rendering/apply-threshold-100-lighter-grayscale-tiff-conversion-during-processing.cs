using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Define input and output directories (adjust the paths to your environment)
        string dataDir = Path.Combine(Environment.CurrentDirectory, "Data");
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Load the source Word document
        Document doc = new Document(Path.Combine(dataDir, "Source.docx"));

        // Configure image save options for TIFF output
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
            // Apply a lower threshold (100) to obtain a lighter grayscale result
            ThresholdForFloydSteinbergDithering = 100,
            // Choose a CCITT compression scheme suitable for bi‑level images
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a multi‑frame TIFF file using the configured options
        doc.Save(Path.Combine(outputDir, "Document.GrayscaleLighter.tiff"), tiffOptions);
    }
}
