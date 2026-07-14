using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Set up input and output folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputHtml");
        string outputDir = Path.Combine(baseDir, "OutputPdf");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files.
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample1.html"),
            "<html><body><h1>Sample 1</h1><p>This is the first sample.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputDir, "sample2.html"),
            "<html><body><h2>Sample 2</h2><p>Second sample content.</p></body></html>");

        // Convert each HTML file to PDF with a custom page size.
        foreach (string htmlPath in Directory.GetFiles(inputDir, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Define a custom page size (A5: 420 x 595 points).
            doc.FirstSection.PageSetup.PageWidth = 420;
            doc.FirstSection.PageSetup.PageHeight = 595;

            // Prepare PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }
    }

    private static void CreateSampleHtmlFile(string path, string htmlContent)
    {
        File.WriteAllText(path, htmlContent);
    }
}
