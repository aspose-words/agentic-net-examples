using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class RtfToPdfConverter
{
    static void Main()
    {
        // Path to the source RTF file.
        string rtfPath = @"C:\Docs\source.rtf";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\result.pdf";

        // Load the RTF document with default load options.
        RtfLoadOptions loadOptions = new RtfLoadOptions();
        Document doc = new Document(rtfPath, loadOptions);

        // Prepare PDF save options (default settings are sufficient for a basic conversion).
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the loaded document as PDF using the specified save options.
        doc.Save(pdfPath, pdfOptions);
    }
}
