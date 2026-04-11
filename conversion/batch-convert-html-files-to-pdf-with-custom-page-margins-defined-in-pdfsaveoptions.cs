using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create a few sample HTML files.
        for (int i = 1; i <= 3; i++)
        {
            string htmlPath = Path.Combine(inputFolder, $"Sample{i}.html");
            string htmlContent = $"<html><body><h1>Document {i}</h1><p>This is a sample HTML file number {i}.</p></body></html>";
            File.WriteAllText(htmlPath, htmlContent);
        }

        // Define custom page margins (in points). 1 point = 1/72 inch.
        // Here we use 20 mm top/bottom and 15 mm left/right margins.
        double topBottomMargin = ConvertUtil.MillimeterToPoint(20);
        double leftRightMargin = ConvertUtil.MillimeterToPoint(15);

        // Process each HTML file in the input folder.
        foreach (string htmlFile in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Apply custom margins to the first (and only) section.
            if (doc.Sections.Count > 0)
            {
                doc.Sections[0].PageSetup.TopMargin = topBottomMargin;
                doc.Sections[0].PageSetup.BottomMargin = topBottomMargin;
                doc.Sections[0].PageSetup.LeftMargin = leftRightMargin;
                doc.Sections[0].PageSetup.RightMargin = leftRightMargin;
            }

            // Prepare PDF save options (no special options needed beyond defaults).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
