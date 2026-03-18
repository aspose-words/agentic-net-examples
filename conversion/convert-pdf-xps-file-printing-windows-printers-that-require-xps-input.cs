using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToXpsConversion
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Input\sample.pdf";

            // Path where the resulting XPS file will be saved.
            string xpsPath = @"C:\Output\sample.xps";

            // Load the PDF document.
            // Uses the Document(string) constructor rule.
            Document pdfDocument = new Document(pdfPath);

            // Create XpsSaveOptions instance.
            // Uses the XpsSaveOptions() constructor rule.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // (Optional) Set any additional options here, e.g. optimize output.
            // xpsOptions.OptimizeOutput = true;

            // Save the document as XPS using the Save(string, SaveOptions) method rule.
            pdfDocument.Save(xpsPath, xpsOptions);

            Console.WriteLine("PDF has been successfully converted to XPS.");
        }
    }
}
