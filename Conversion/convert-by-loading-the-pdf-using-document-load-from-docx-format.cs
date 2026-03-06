using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Load the PDF using PdfLoadOptions (inherits from LoadOptions).
        PdfLoadOptions loadOptions = new PdfLoadOptions();

        // Load the PDF document into an Aspose.Words Document object.
        Document doc = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX.
        string docxPath = @"C:\Output\sample.docx";
        doc.Save(docxPath, SaveFormat.Docx);
    }
}
