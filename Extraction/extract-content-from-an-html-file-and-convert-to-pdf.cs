using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HtmlToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file.
            string htmlFilePath = @"C:\Input\sample.html";

            // Path where the resulting PDF will be saved.
            string pdfFilePath = @"C:\Output\sample.pdf";

            // Load the HTML document. The constructor automatically detects the format.
            Document doc = new Document(htmlFilePath);

            // Save the document as PDF. The Save method determines the format from the file extension.
            doc.Save(pdfFilePath, SaveFormat.Pdf);

            // Optional: inform the user that the conversion has completed.
            Console.WriteLine("HTML file has been successfully converted to PDF.");
        }
    }
}
