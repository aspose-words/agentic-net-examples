using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace HtmlToPdfConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file.
            string htmlFilePath = @"C:\Temp\source.html";

            // Path where the resulting PDF will be saved.
            string pdfFilePath = @"C:\Temp\result.pdf";

            // Load the HTML document using HtmlLoadOptions (default settings).
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Document document = new Document(htmlFilePath, loadOptions);

            // Prepare PDF save options (default settings).
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Save the loaded document as PDF.
            document.Save(pdfFilePath, saveOptions);
        }
    }
}
