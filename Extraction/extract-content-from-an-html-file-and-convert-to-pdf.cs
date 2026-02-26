using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HtmlToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source HTML file.
            // Replace with the actual file location as needed.
            string htmlFilePath = @"C:\Path\To\Your\Input.html";

            // Path where the resulting PDF will be saved.
            // The file extension determines the output format.
            string pdfFilePath = @"C:\Path\To\Your\Output.pdf";

            // Load the HTML document. The Document constructor automatically detects the format.
            Document document = new Document(htmlFilePath);

            // Save the loaded document as PDF.
            // The Save method overload with (string, SaveFormat) explicitly sets the target format.
            document.Save(pdfFilePath, SaveFormat.Pdf);
        }
    }
}
