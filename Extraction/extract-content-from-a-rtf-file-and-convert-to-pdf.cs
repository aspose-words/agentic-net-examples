using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace RtfToPdfExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source RTF file.
            string rtfPath = @"C:\Input\sample.rtf";

            // Path where the resulting PDF will be saved.
            string pdfPath = @"C:\Output\sample.pdf";

            // Load the RTF document using RtfLoadOptions (default options are sufficient here).
            RtfLoadOptions loadOptions = new RtfLoadOptions();
            Document doc = new Document(rtfPath, loadOptions);

            // Prepare PDF save options (default options are sufficient for a basic conversion).
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Save the loaded document as PDF.
            doc.Save(pdfPath, pdfSaveOptions);
        }
    }
}
