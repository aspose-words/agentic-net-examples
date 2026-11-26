using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document saved as compressed PDF/A‑2u.");

        // Configure PDF/A‑2u save options.
        // Enable JPEG image compression with a quality setting.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 80
            // The older properties RemoveUnusedObjects and CompressContentStreams are no longer
            // available in recent Aspose.Words versions; the default behavior already optimizes the PDF.
        };

        // Save the document to a PDF/A‑2u file.
        string outputPath = "compressed_pdfa2u.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created and is not empty.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");
        }

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
        {
            throw new InvalidOperationException("The output PDF file is empty.");
        }

        Console.WriteLine($"PDF/A‑2u file saved successfully: {outputPath} (size: {info.Length} bytes)");
    }
}
