using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Configure TXT save options to include headers and footers.
        TxtSaveOptions txtOptions = new TxtSaveOptions
        {
            // Export only the primary header/footer at the start/end of each section.
            ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly
        };

        // Save the document as plain text, preserving the configured headers/footers.
        doc.Save("OutputDocument.txt", txtOptions);
    }
}
