using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDFs.
        string inputFolder = Path.Combine(Directory.GetCurrentDirectory(), "InputHtml");
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputPdf");

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files if the folder is empty.
        if (Directory.GetFiles(inputFolder, "*.html").Length == 0)
        {
            File.WriteAllText(Path.Combine(inputFolder, "Sample1.html"),
                "<html><body><h1>Sample 1</h1><p>This is the first sample HTML file.</p></body></html>");
            File.WriteAllText(Path.Combine(inputFolder, "Sample2.html"),
                "<html><body><h1>Sample 2</h1><p>This is the second sample HTML file.</p></body></html>");
        }

        // Process each HTML file in the input folder.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Define a custom page size (width x height in points).
            // Example: 500 points wide by 700 points high.
            doc.FirstSection.PageSetup.PageWidth = 500;
            doc.FirstSection.PageSetup.PageHeight = 700;

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF file name.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the custom options.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Optional: indicate completion (no interactive input required).
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
