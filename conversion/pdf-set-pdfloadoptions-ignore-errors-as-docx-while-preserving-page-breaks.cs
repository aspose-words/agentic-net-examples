using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToDocxConverter
{
    static void Main()
    {
        // Input PDF file path.
        string pdfPath = "input.pdf";

        // Output DOCX file path.
        string docxPath = "output.docx";

        // Create load options for PDF and set recovery mode to try to recover
        // (ignore errors that may occur while loading the PDF).
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.RecoveryMode = DocumentRecoveryMode.TryRecover;

        // Load the PDF document using the load options.
        Document document = new Document(pdfPath, loadOptions);

        // Save the loaded document as DOCX. Page breaks are part of the document
        // structure and are preserved automatically when saving to DOCX.
        document.Save(docxPath, SaveFormat.Docx);
    }
}
