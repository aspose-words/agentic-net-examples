using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Input DOCX file and output PDF file paths.
        string inputPath = @"C:\Docs\Sample.docx";
        string outputPath = @"C:\Docs\Sample.pdf";

        // Load the DOCX document from the file system.
        Document doc = new Document(inputPath);

        // Create PDF save options and configure custom settings.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Render all images in grayscale.
            ColorMode = ColorMode.Grayscale,
            // Use high‑quality (slower) rendering algorithms.
            UseHighQualityRendering = true,
            // Optimize memory usage for large documents.
            MemoryOptimization = true
        };

        // Configure outline (bookmarks) options: include only the first three heading levels
        // and expand only the top‑level entries when the PDF is opened.
        pdfOptions.OutlineOptions.HeadingsOutlineLevels = 3;
        pdfOptions.OutlineOptions.ExpandedOutlineLevels = 1;

        // Save the document as PDF using the custom options.
        doc.Save(outputPath, pdfOptions);
    }
}
