using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPngBatchConverter
{
    // Converts all PDF files in the specified input folder to high‑resolution PNG images (600 dpi).
    // Each page of a PDF is saved as a separate PNG file named "<pdfFileName>_Page<pageNumber>.png".
    static void Main()
    {
        // Folder that contains the source PDF files.
        string inputFolder = @"C:\InputPdfs";

        // Folder where the PNG images will be written.
        string outputFolder = @"C:\OutputPngs";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Get all PDF files in the input folder.
        string[] pdfFiles = Directory.GetFiles(inputFolder, "*.pdf", SearchOption.TopDirectoryOnly);

        // Prepare the image save options once – they will be reused for every page.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set both horizontal and vertical resolution to 600 dpi.
            Resolution = 600f,

            // Optional: improve rendering quality (slower but higher fidelity).
            UseAntiAliasing = true,
            UseHighQualityRendering = true
        };

        foreach (string pdfPath in pdfFiles)
        {
            // Load the PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Iterate through all pages of the current PDF.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                // Configure the PageSet to render only the current page.
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name: "MyFile_Page1.png", "MyFile_Page2.png", etc.
                string pdfFileName = Path.GetFileNameWithoutExtension(pdfPath);
                string outputFile = Path.Combine(
                    outputFolder,
                    $"{pdfFileName}_Page{pageIndex + 1}.png");

                // Save the selected page as a PNG image using the prepared options.
                pdfDoc.Save(outputFile, pngOptions);
            }
        }

        Console.WriteLine("Conversion completed.");
    }
}
