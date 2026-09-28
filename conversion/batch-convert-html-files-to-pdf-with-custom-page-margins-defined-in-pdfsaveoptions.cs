using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input HTML files and output PDF files.
        string inputFolder = "input_html";
        string outputFolder = "output_pdf";

        // Ensure the folders exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);

        // Create sample HTML files.
        string htmlFile1 = Path.Combine(inputFolder, "sample1.html");
        string htmlContent1 = "<html><body><h1>Sample 1</h1><p>This is the first sample HTML file.</p></body></html>";
        File.WriteAllText(htmlFile1, htmlContent1);

        string htmlFile2 = Path.Combine(inputFolder, "sample2.html");
        string htmlContent2 = "<html><body><h2>Sample 2</h2><p>This is the second sample HTML file.</p></body></html>";
        File.WriteAllText(htmlFile2, htmlContent2);

        // Define custom margins (in points; 72 points = 1 inch).
        const double marginTop = 72;    // 1 inch
        const double marginBottom = 72; // 1 inch
        const double marginLeft = 72;   // 1 inch
        const double marginRight = 72;  // 1 inch

        // Process each HTML file in the input folder.
        foreach (string htmlPath in Directory.GetFiles(inputFolder, "*.html"))
        {
            // Load the HTML document.
            Document doc = new Document(htmlPath);

            // Apply custom margins to every section in the document.
            foreach (Section section in doc.Sections)
            {
                PageSetup setup = section.PageSetup;
                setup.TopMargin = marginTop;
                setup.BottomMargin = marginBottom;
                setup.LeftMargin = marginLeft;
                setup.RightMargin = marginRight;
            }

            // Configure PDF save options (no special page info needed).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Determine the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the specified options.
            doc.Save(pdfPath, pdfOptions);

            // Validate that the PDF file was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Expected PDF file was not created: {pdfPath}");
        }

        // Indicate successful completion.
        Console.WriteLine("Batch conversion completed successfully.");
    }
}
