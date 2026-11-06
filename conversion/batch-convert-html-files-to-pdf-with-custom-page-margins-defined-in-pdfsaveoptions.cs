using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare input and output folders
        string inputFolder = "HtmlInput";
        string outputFolder = "PdfOutput";

        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files
        string html1Path = Path.Combine(inputFolder, "sample1.html");
        string html2Path = Path.Combine(inputFolder, "sample2.html");

        File.WriteAllText(html1Path,
            "<html><body><h1>First Document</h1><p>This is the first sample HTML.</p></body></html>");
        File.WriteAllText(html2Path,
            "<html><body><h2>Second Document</h2><p>This is the second sample HTML.</p></body></html>");

        // Define custom margins (in points; 1 inch = 72 points)
        const double leftMargin = 72;   // 1 inch
        const double rightMargin = 72;  // 1 inch
        const double topMargin = 72;    // 1 inch
        const double bottomMargin = 72; // 1 inch

        // Batch convert each HTML file to PDF with the custom margins
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load HTML document
            Document doc = new Document(htmlPath);

            // Apply custom margins to every section in the document
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.LeftMargin = leftMargin;
                section.PageSetup.RightMargin = rightMargin;
                section.PageSetup.TopMargin = topMargin;
                section.PageSetup.BottomMargin = bottomMargin;
            }

            // Use default PDF save options (no special options required for margins)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save as PDF
            doc.Save(pdfPath, pdfOptions);

            // Validate that the PDF was created
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"PDF file was not created: {pdfPath}");
        }

        // Final validation: ensure at least one PDF exists
        string[] pdfFiles = Directory.GetFiles(outputFolder, "*.pdf");
        if (pdfFiles.Length == 0)
            throw new InvalidOperationException("No PDF files were generated in the output folder.");

        Console.WriteLine("Batch HTML to PDF conversion completed successfully.");
    }
}
