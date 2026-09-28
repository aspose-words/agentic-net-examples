using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample document and save it as PDF (input.pdf).
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("This is a sample PDF document generated for conversion to PNG.");
        string inputPdfPath = "input.pdf";
        sampleDoc.Save(inputPdfPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(inputPdfPath))
            throw new InvalidOperationException("The input PDF file was not created.");

        // Step 2: Load the PDF document.
        Document pdfDoc = new Document(inputPdfPath);

        // Step 3: Configure high‑resolution PNG conversion.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set a high DPI for detailed analysis (e.g., 300 DPI).
            Resolution = 300,
            // Save only the first page (page index is zero‑based).
            PageSet = new PageSet(0)
        };

        // Step 4: Save the PDF as a PNG image.
        string outputPngPath = "output.png";
        pdfDoc.Save(outputPngPath, pngOptions);

        // Step 5: Validate that the PNG file was created.
        if (!File.Exists(outputPngPath))
            throw new InvalidOperationException("The output PNG file was not created.");

        // Optional: Clean up temporary files (comment out if you want to keep them).
        // File.Delete(inputPdfPath);
        // File.Delete(outputPngPath);
    }
}
