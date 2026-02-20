using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class AddWatermarkToPdfWithFormFields
{
    static void Main()
    {
        // Load the source Word document that contains form fields.
        Document doc = new Document("InputWithFormFields.docx");

        // Configure text watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.Red,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Add the text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Set PDF save options to preserve form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };

        // Save the document as PDF with the watermark and preserved form fields.
        doc.Save("OutputWithWatermark.pdf", pdfOptions);
    }
}
