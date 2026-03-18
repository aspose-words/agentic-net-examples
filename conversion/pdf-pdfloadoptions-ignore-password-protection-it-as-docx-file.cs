using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file (may be password‑protected).
        const string pdfPath = @"C:\Temp\source.pdf";

        // Desired output DOCX file path.
        const string docxPath = @"C:\Temp\converted.docx";

        // Create PdfLoadOptions. By leaving Password null (default) we tell Aspose.Words
        // to ignore any password protection on the PDF.
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Load the PDF into an Aspose.Words Document using the load options.
        Document doc = new Document(pdfPath, loadOptions);

        // Save the document as DOCX. The SaveFormat enum ensures the correct format.
        doc.Save(docxPath, SaveFormat.Docx);
    }
}
