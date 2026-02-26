using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class RtfToPdfConverter
{
    static void Main()
    {
        // Input RTF file path
        string rtfPath = "input.rtf";

        // Output PDF file path
        string pdfPath = "output.pdf";

        // Load the RTF document using the default RtfLoadOptions
        Document doc = new Document(rtfPath, new RtfLoadOptions());

        // Save the loaded document as PDF using the default PdfSaveOptions
        doc.Save(pdfPath, new PdfSaveOptions());
    }
}
