using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfThumbnailBatch
{
    static void Main()
    {
        // Folder containing source PDF files
        string inputFolder = @"C:\InputPdfs";

        // Folder where JPEG thumbnails will be saved
        string outputFolder = @"C:\Thumbnails";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each PDF file in the input folder
        foreach (string pdfPath in Directory.GetFiles(inputFolder, "*.pdf"))
        {
            // Load the PDF document
            Document doc = new Document(pdfPath);

            // Set up image save options for JPEG with low quality
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);
            // Render only the first page (zero‑based index)
            options.PageSet = new PageSet(0);
            // Low JPEG quality for strong compression (e.g., 10)
            options.JpegQuality = 10;

            // Construct the output file name
            string thumbFileName = Path.GetFileNameWithoutExtension(pdfPath) + "_thumb.jpg";
            string thumbPath = Path.Combine(outputFolder, thumbFileName);

            // Save the first‑page thumbnail as JPEG
            doc.Save(thumbPath, options);
        }
    }
}
