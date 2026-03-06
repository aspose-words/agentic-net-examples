using System;
using Aspose.Words;
using Aspose.Words.Loading;

namespace HtmlToPdfSample
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file.
            string htmlFile = @"C:\Temp\sample.html";

            // Path where the resulting PDF will be saved.
            string pdfFile = @"C:\Temp\sample.pdf";

            // Load the HTML document. HtmlLoadOptions can be omitted if default options are sufficient.
            Document doc = new Document(htmlFile, new HtmlLoadOptions());

            // Save the loaded document as PDF. The SaveFormat enum specifies the output format.
            doc.Save(pdfFile, SaveFormat.Pdf);
        }
    }
}
