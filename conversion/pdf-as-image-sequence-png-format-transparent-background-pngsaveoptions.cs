using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPngSequence
{
    static void Main()
    {
        // Create a temporary PDF file to work with.
        string tempPdfPath = Path.Combine(Path.GetTempPath(), "sample.pdf");
        if (!File.Exists(tempPdfPath))
        {
            Document tempDoc = new Document();
            var builder = new DocumentBuilder(tempDoc);
            builder.Writeln("This is a sample PDF generated for the demo.");
            tempDoc.Save(tempPdfPath, SaveFormat.Pdf);
        }

        // Folder where the PNG images will be saved (relative to the current directory).
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "OutputImages");
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document pdfDoc = new Document(tempPdfPath);

        // Create ImageSaveOptions for PNG format with a transparent background.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PaperColor = Color.Transparent
        };

        // Iterate through each page of the PDF and save as PNG.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            pngOptions.PageSet = new PageSet(pageIndex);
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            pdfDoc.Save(outputPath, pngOptions);
        }

        Console.WriteLine("PDF pages have been saved as PNG images with transparent background.");
    }
}
