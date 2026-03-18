using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class HtmlToPdfBatchConverter
{
    // Custom margin in millimeters (applied to all sides)
    const double MarginInMillimeters = 15.0;

    static void Main()
    {
        // Folder containing the source HTML files
        string sourceFolder = @"C:\InputHtml";

        // Folder where the resulting PDF files will be saved
        string outputFolder = @"C:\OutputPdf";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Convert each .html file in the source folder
        foreach (string htmlFilePath in Directory.GetFiles(sourceFolder, "*.html"))
        {
            // Load the HTML document
            Document doc = new Document(htmlFilePath);

            // Apply custom margins to every section of the document
            double marginPoints = ConvertUtil.MillimeterToPoint(MarginInMillimeters);
            foreach (Section section in doc.Sections)
            {
                section.PageSetup.TopMargin = marginPoints;
                section.PageSetup.BottomMargin = marginPoints;
                section.PageSetup.LeftMargin = marginPoints;
                section.PageSetup.RightMargin = marginPoints;
            }

            // Prepare PDF save options (no special options required for margins)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Build the output PDF file name (same base name as the HTML file)
            string pdfFileName = Path.GetFileNameWithoutExtension(htmlFilePath) + ".pdf";
            string pdfFilePath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF using the specified options
            doc.Save(pdfFilePath, pdfOptions);
        }

        Console.WriteLine("Batch conversion completed.");
    }
}
