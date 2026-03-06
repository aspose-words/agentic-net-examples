using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfConversionExampleApp
{
    public class PdfConversionExample
    {
        // Converts a PDF file to DOCX, HTML and renders each page as a PNG image.
        public static void ConvertPdf(string pdfPath, string outputFolder)
        {
            // Load the source PDF document.
            Document pdfDoc = new Document(pdfPath);

            // Save as DOCX.
            pdfDoc.Save($"{outputFolder}/Converted.docx", SaveFormat.Docx);

            // Save as HTML.
            pdfDoc.Save($"{outputFolder}/Converted.html", SaveFormat.Html);

            // Prepare image save options for PNG rendering.
            ImageSaveOptions imgOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                Resolution = 300,                         // Set DPI.
                ImageSize = new Size(1240, 1754)          // Optional: define image dimensions.
            };

            // Render each page of the PDF to a separate PNG file.
            for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
            {
                imgOptions.PageSet = new PageSet(pageIndex);
                pdfDoc.Save($"{outputFolder}/Page_{pageIndex + 1}.png", imgOptions);
            }
        }
    }

    class Program
    {
        // Entry point required for a console application.
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: PdfConversionExample <pdfPath> <outputFolder>");
                return;
            }

            string pdfPath = args[0];
            string outputFolder = args[1];

            try
            {
                PdfConversionExample.ConvertPdf(pdfPath, outputFolder);
                Console.WriteLine("PDF conversion completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
