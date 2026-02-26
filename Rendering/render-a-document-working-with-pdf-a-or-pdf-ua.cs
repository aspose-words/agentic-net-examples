using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Load an existing Word document.
        Document doc = new Document("Input.docx");

        // Create PDF save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the compliance level to PDF/A-2u (ISO 19005-2) which preserves visual appearance
        // and allows reliable text extraction. Change to any other PdfCompliance value as needed.
        saveOptions.Compliance = PdfCompliance.PdfA2u;

        // Save the document as a PDF that conforms to the selected standard.
        doc.Save("Output_PdfA2u.pdf", saveOptions);
    }
}
