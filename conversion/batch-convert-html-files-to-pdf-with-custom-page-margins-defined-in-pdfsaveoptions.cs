using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToPdfConverter
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        CreateSampleHtmlFile(Path.Combine(inputFolder, "sample1.html"),
            "<html><body><h1>Sample 1</h1><p>This is the first sample.</p></body></html>");
        CreateSampleHtmlFile(Path.Combine(inputFolder, "sample2.html"),
            "<html><body><h1>Sample 2</h1><p>This is the second sample.</p></body></html>");

        // Configure PDF save options (no special page info needed here).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Process each HTML file in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlPath in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Set custom page margins (1 inch = 72 points) for every section.
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.TopMargin = 72;
                section.PageSetup.BottomMargin = 72;
                section.PageSetup.LeftMargin = 72;
                section.PageSetup.RightMargin = 72;
            }

            // Determine the output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the custom options.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        // Verify that at least one PDF was generated.
        if (Directory.GetFiles(outputFolder, "*.pdf").Length == 0)
            throw new InvalidOperationException("No PDF files were generated in the output folder.");
    }

    private static void CreateSampleHtmlFile(string path, string htmlContent)
    {
        File.WriteAllText(path, htmlContent);
        if (!File.Exists(path))
            throw new InvalidOperationException($"Failed to create sample HTML file: {path}");
    }
}
