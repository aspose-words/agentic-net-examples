using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExportHeaderFooterToTxt
{
    class Program
    {
        static void Main()
        {
            // Load the existing DOCX document.
            Document doc = new Document("InputDocument.docx");

            // Configure TXT save options to include both primary and even headers/footers.
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                // Export all headers and footers at the end of the document.
                ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd
            };

            // Save the document as plain text using the configured options.
            doc.Save("OutputDocument.txt", txtOptions);
        }
    }
}
