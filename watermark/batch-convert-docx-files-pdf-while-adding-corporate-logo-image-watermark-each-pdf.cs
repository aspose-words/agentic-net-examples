using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;   // For ImageWatermarkOptions

class BatchDocxToPdfWithWatermark
{
    static void Main()
    {
        // Folder that contains the source DOCX files.
        string inputFolder = @"C:\Docs\Input";

        // Folder where the resulting PDF files will be written.
        string outputFolder = @"C:\Docs\Output";

        // Path to the corporate logo that will be used as an image watermark.
        string logoPath = @"C:\Images\CorporateLogo.png";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Configure the watermark appearance (optional).
        ImageWatermarkOptions watermarkOptions = new ImageWatermarkOptions
        {
            // Scale the logo to 30 % of the page width (adjust as needed).
            Scale = 0.3,
            // Do not wash out the image; keep original colors.
            IsWashout = false
        };

        // Process each DOCX file in the input folder.
        foreach (string docxFile in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document (lifecycle rule: load).
            Document document = new Document(docxFile);

            // Apply the corporate logo as an image watermark (feature rule).
            document.Watermark.SetImage(logoPath, watermarkOptions);

            // Determine the output PDF file name.
            string pdfFile = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(docxFile) + ".pdf");

            // Save the document as PDF (lifecycle rule: save).
            document.Save(pdfFile);
        }
    }
}
