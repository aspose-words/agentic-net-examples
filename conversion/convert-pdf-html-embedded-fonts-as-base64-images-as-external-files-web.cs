using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToHtmlConversion
{
    public class Converter
    {
        /// <summary>
        /// Converts a PDF document to HTML.
        /// Fonts are embedded in the HTML as Base64 strings, while images are saved as external files.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="outputHtmlPath">Full path where the resulting HTML file will be saved.</param>
        /// <param name="imagesFolder">Folder that will contain the extracted image files.</param>
        public static void ConvertPdfToHtml(string pdfPath, string outputHtmlPath, string imagesFolder)
        {
            // Ensure the images folder exists.
            if (!Directory.Exists(imagesFolder))
                Directory.CreateDirectory(imagesFolder);

            // Load the PDF document.
            Document doc = new Document(pdfPath);

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Embed fonts directly into the CSS using Base64 encoding.
                ExportFontsAsBase64 = true,

                // Use an embedded stylesheet so that the Base64 font data is placed in the HTML.
                CssStyleSheetType = CssStyleSheetType.Embedded,

                // Write the HTML in a readable (pretty) format.
                PrettyFormat = true,

                // Save images as separate files (default behavior), specifying the folder for them.
                ImagesFolder = imagesFolder,

                // Ensure images are NOT embedded as Base64.
                ExportImagesAsBase64 = false
            };

            // Save the document as HTML using the configured options.
            doc.Save(outputHtmlPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            string pdfFile = @"C:\Input\sample.pdf";
            string htmlFile = @"C:\Output\sample.html";
            string imagesDir = @"C:\Output\sample_images";

            ConvertPdfToHtml(pdfFile, htmlFile, imagesDir);
            Console.WriteLine("Conversion completed.");
        }
    }
}
