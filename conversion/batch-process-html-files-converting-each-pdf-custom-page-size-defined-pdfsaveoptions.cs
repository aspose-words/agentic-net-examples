using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class HtmlToPdfBatch
{
    static void Main()
    {
        // Folder containing the source HTML files (relative to the executable).
        string inputFolder = Path.Combine(AppContext.BaseDirectory, "InputHtml");

        // Folder where the resulting PDF files will be saved (relative to the executable).
        string outputFolder = Path.Combine(AppContext.BaseDirectory, "OutputPdf");

        // Ensure both directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no HTML files, create a simple sample file so the demo can run.
        if (Directory.GetFiles(inputFolder, "*.html").Length == 0)
        {
            string sampleHtmlPath = Path.Combine(inputFolder, "Sample.html");
            File.WriteAllText(sampleHtmlPath, "<html><body><h1>Hello, Aspose.Words!</h1></body></html>");
        }

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
                section.PageSetup.PaperSize = PaperSize.Custom;
                section.PageSetup.PageWidth = customPageWidth;
                section.PageSetup.PageHeight = customPageHeight;
            }

            // Create PdfSaveOptions to control PDF conversion.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the specified options.
            doc.Save(pdfPath, pdfOptions);
        }

        Console.WriteLine($"Conversion completed. PDFs are saved in: {outputFolder}");
    }
}
