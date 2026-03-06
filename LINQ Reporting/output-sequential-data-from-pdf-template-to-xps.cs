using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfToXpsConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF template.
            string pdfTemplatePath = @"C:\Input\Template.pdf";

            // Path where the resulting XPS file will be saved.
            string xpsOutputPath = @"C:\Output\Result.xps";

            // Load the PDF document. Aspose.Words can open PDF files directly.
            Document document = new Document(pdfTemplatePath);

            // Create XpsSaveOptions to control the XPS conversion.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();

            // Example: ensure the document generator name is embedded (default is true).
            xpsOptions.ExportGeneratorName = true;

            // Save the loaded document as XPS using the specified options.
            document.Save(xpsOutputPath, xpsOptions);
        }
    }
}
