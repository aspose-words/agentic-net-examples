using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define input and output directories.
        string inputDir = "InputHtml";
        string outputDir = "OutputPdf";

        // Ensure the directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create sample HTML files.
        File.WriteAllText(Path.Combine(inputDir, "sample1.html"),
            "<html><body><h1>Sample 1</h1><p>Hello World!</p></body></html>");
        File.WriteAllText(Path.Combine(inputDir, "sample2.html"),
            "<html><body><h1>Sample 2</h1><p>Another page.</p></body></html>");

        // Get all HTML files in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputDir, "*.html");

        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Set custom page margins (72 points = 1 inch).
            doc.FirstSection.PageSetup.TopMargin = 72;
            doc.FirstSection.PageSetup.BottomMargin = 72;
            doc.FirstSection.PageSetup.LeftMargin = 72;
            doc.FirstSection.PageSetup.RightMargin = 72;

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

        // Ensure at least one PDF was generated.
        if (Directory.GetFiles(outputDir, "*.pdf").Length == 0)
            throw new InvalidOperationException("No PDF files were generated.");
    }
}
