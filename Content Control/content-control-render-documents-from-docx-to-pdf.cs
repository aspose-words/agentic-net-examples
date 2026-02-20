using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file.
            string sourcePath = "input.docx";

            // Path where the resulting PDF will be saved.
            string destinationPath = "output.pdf";

            // Load the DOCX document.
            Document document = new Document(sourcePath);

            // Create PDF save options (default settings).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the document as PDF using the specified options.
            document.Save(destinationPath, pdfOptions);
        }
    }
}
