using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "HtmlInputs");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "PdfOutputs");

        // Ensure clean environment.
        if (Directory.Exists(inputFolder))
            Directory.Delete(inputFolder, true);
        if (Directory.Exists(outputFolder))
            Directory.Delete(outputFolder, true);
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        CreateSampleHtml(Path.Combine(inputFolder, "Sample1.html"), "<html><body><h1>First Document</h1><p>Hello from HTML 1.</p></body></html>");
        CreateSampleHtml(Path.Combine(inputFolder, "Sample2.html"), "<html><body><h1>Second Document</h1><p>Hello from HTML 2.</p></body></html>");

        // Define custom margins (in points). 1 inch = 72 points.
        const double leftMargin = 72;   // 1 inch
        const double rightMargin = 72;  // 1 inch
        const double topMargin = 72;    // 1 inch
        const double bottomMargin = 72; // 1 inch

        // Prepare PdfSaveOptions (no specific margin property exists, margins are set on the document).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Process each HTML file.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Apply custom page margins to each section.
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.LeftMargin = leftMargin;
                section.PageSetup.RightMargin = rightMargin;
                section.PageSetup.TopMargin = topMargin;
                section.PageSetup.BottomMargin = bottomMargin;
            }

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF using the options.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // All conversions completed successfully.
        Console.WriteLine("Batch conversion completed. PDFs are located in: " + outputFolder);
    }

    private static void CreateSampleHtml(string filePath, string htmlContent)
    {
        File.WriteAllText(filePath, htmlContent);
    }
}
