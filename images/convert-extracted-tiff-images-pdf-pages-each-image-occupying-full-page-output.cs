using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class TiffToPdfConverter
{
    static void Main()
    {
        // Create a temporary folder for the generated images and the output PDF.
        string tempFolder = Path.Combine(Path.GetTempPath(), "TiffToPdfExample");
        Directory.CreateDirectory(tempFolder);

        // Generate sample PNG images (a 1x1 pixel image) to simulate TIFF pages.
        string[] imageFiles = new string[3];
        byte[] pngData = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X6V8AAAAASUVORK5CYII=");

        for (int i = 0; i < imageFiles.Length; i++)
        {
            string filePath = Path.Combine(tempFolder, $"Page{i + 1}.png");
            File.WriteAllBytes(filePath, pngData);
            imageFiles[i] = filePath;
        }

        // Create a new empty Word document – this will become the PDF container.
        Document pdfDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(pdfDoc);

        // Insert each image as a full‑page picture.
        for (int i = 0; i < imageFiles.Length; i++)
        {
            builder.InsertImage(imageFiles[i]);

            // After each image except the last, insert a page break so the next image starts on a new page.
            if (i < imageFiles.Length - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the assembled document as a PDF file.
        string outputPdf = Path.Combine(tempFolder, "Combined.pdf");
        pdfDoc.Save(outputPdf, SaveFormat.Pdf);

        Console.WriteLine($"PDF created successfully at: {outputPdf}");
    }
}
