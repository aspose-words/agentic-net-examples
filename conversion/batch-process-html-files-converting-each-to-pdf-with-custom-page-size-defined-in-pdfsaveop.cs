using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class BatchHtmlToPdf
{
    public static void Main()
    {
        // Define input and output folders.
        string inputFolder = "InputHtml";
        string outputFolder = "OutputPdf";

        // Ensure folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        string html1Path = Path.Combine(inputFolder, "sample1.html");
        string html2Path = Path.Combine(inputFolder, "sample2.html");

        File.WriteAllText(html1Path,
            "<html><body><h1>Sample 1</h1><p>This is the first HTML file.</p></body></html>");
        File.WriteAllText(html2Path,
            "<html><body><h1>Sample 2</h1><p>This is the second HTML file.</p></body></html>");

        // Process each HTML file in the input folder.
        string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");
        foreach (string htmlFile in htmlFiles)
        {
            // Load the HTML document.
            Document doc = new Document(htmlFile);

            // Set a custom page size (A5: 420x595 points) for the first section.
            // Aspose.Words uses points for page dimensions.
            Section firstSection = doc.FirstSection;
            if (firstSection != null)
            {
                PageSetup pageSetup = firstSection.PageSetup;
                pageSetup.PageWidth = 420;   // Width in points.
                pageSetup.PageHeight = 595;  // Height in points.
            }

            // Determine output PDF path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlFile) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF file was not created: {pdfPath}");
        }

        // Indicate completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
