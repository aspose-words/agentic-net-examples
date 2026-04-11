using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define paths for the temporary files.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        string pdfPath = Path.Combine(artifactsDir, "sample.pdf");
        string jpegPath = Path.Combine(artifactsDir, "sample.jpg");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as PDF (input file).
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF that will be converted to a high‑resolution JPEG image.");
        // Save the document as PDF.
        doc.Save(pdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 2. Load the PDF document.
        // -----------------------------------------------------------------
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // 3. Configure image save options for high resolution and high quality.
        // -----------------------------------------------------------------
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // High resolution (dots per inch). Adjust as needed.
            Resolution = 300f,
            // JPEG quality: 100 = best quality, minimal compression.
            JpegQuality = 100
        };

        // -----------------------------------------------------------------
        // 4. Save the first page of the PDF as a JPEG image.
        // -----------------------------------------------------------------
        pdfDoc.Save(jpegPath, imageOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created and is not empty.
        // -----------------------------------------------------------------
        if (!File.Exists(jpegPath))
            throw new FileNotFoundException("The JPEG output file was not created.", jpegPath);

        FileInfo info = new FileInfo(jpegPath);
        if (info.Length == 0)
            throw new InvalidOperationException("The JPEG output file is empty.");

        // Indicate successful completion.
        Console.WriteLine($"PDF successfully converted to high‑resolution JPEG: {jpegPath}");
    }
}
