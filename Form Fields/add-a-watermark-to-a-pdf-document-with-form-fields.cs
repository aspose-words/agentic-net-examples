using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source Word document that contains form fields.
        Document doc = new Document("Input.docx");

        // Define watermark appearance.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };

        // Apply the text watermark to every page of the document.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Set PDF save options to preserve Word form fields as interactive PDF form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true
        };

        // Save the document as PDF with the watermark and preserved form fields.
        doc.Save("Output.pdf", pdfOptions);
    }
}
