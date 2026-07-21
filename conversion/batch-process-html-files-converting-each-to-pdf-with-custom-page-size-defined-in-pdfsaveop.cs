using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputPdf";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files if the folder is empty.
        if (!Directory.EnumerateFiles(inputFolder, "*.html").Any())
        {
            File.WriteAllText(Path.Combine(inputFolder, "Sample1.html"),
                "<html><body><h1>Sample Document 1</h1><p>This is the first sample.</p></body></html>");
            File.WriteAllText(Path.Combine(inputFolder, "Sample2.html"),
                "<html><body><h1>Sample Document 2</h1><p>This is the second sample.</p></body></html>");
        }

        // Process each HTML file in the input folder.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Define a custom page size (width and height in points).
            // 1 point = 1/72 inch. Example: 500pt x 700pt.
            doc.FirstSection.PageSetup.PageWidth = 500;
            doc.FirstSection.PageSetup.PageHeight = 700;

            // Prepare PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the custom options.
            doc.Save(pdfPath, pdfOptions);

            // Verify that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF file: {pdfPath}");
        }

        // All files processed successfully.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
