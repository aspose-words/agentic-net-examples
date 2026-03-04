using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class WatermarkPdfWithFormFields
{
    static void Main()
    {
        // Load an existing Word document that contains form fields.
        Document doc = new Document("InputDocument.docx");

        // Configure text watermark options.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Prepare PDF save options to preserve Word form fields as interactive PDF form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };

        // Save the document as PDF with the watermark and preserved form fields.
        doc.Save("OutputDocument.pdf", pdfOptions);
    }
}
