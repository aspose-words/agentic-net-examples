using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using Aspose.Words.Rendering;

class HtmlToPdfBatch
{
    static void Main()
    {
        // Folder containing the source HTML files.
        string inputFolder = @"C:\InputHtml";

        // Folder where the resulting PDF files will be saved.
        string outputFolder = @"C:\OutputPdf";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Define custom page dimensions (in points). 1 inch = 72 points.
        // Example: 8.5 x 11 inches (Letter size) but you can change as needed.
        double customPageWidth = 8.5 * 72;   // 612 points
        double customPageHeight = 11 * 72;  // 792 points

        // Process each .html file in the input folder.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Apply custom page size to every section in the document.
            foreach (Section section in doc.Sections)
            {
                // Set the paper size to Custom and specify width/height.
                section.PageSetup.PaperSize = PaperSize.Custom;
                section.PageSetup.PageWidth = customPageWidth;
                section.PageSetup.PageHeight = customPageHeight;
            }

            // Create PdfSaveOptions to control PDF conversion.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // (Optional) Example of setting a PDF-specific option.
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Determine output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the specified options.
            doc.Save(pdfPath, pdfOptions);
        }
    }
}
