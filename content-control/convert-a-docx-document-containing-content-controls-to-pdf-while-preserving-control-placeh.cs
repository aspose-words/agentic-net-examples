using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

namespace ContentControlToPdf
{
    public class Program
    {
        public static void Main()
        {
            // Create a new DOCX document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add introductory text.
            builder.Writeln("Document containing a content control with a placeholder:");

            // Create an inline plain‑text content control (SDT).
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name",
                IsShowingPlaceholderText = true // Show placeholder when the control is empty.
            };

            // Ensure the control has no child nodes so the placeholder is displayed.
            sdt.RemoveAllChildren();

            // Append the content control to the first paragraph.
            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            paragraph.AppendChild(sdt);

            // Save the source DOCX.
            const string inputPath = "input.docx";
            doc.Save(inputPath);

            // Load the DOCX and convert it to PDF, preserving the content control as a form field.
            Document loadedDoc = new Document(inputPath);
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PreserveFormFields = true,          // Keep form fields in the PDF.
                UseSdtTagAsFormFieldName = true    // Use the SDT Tag as the PDF form field name.
            };
            const string outputPath = "output.pdf";
            loadedDoc.Save(outputPath, pdfOptions);
        }
    }
}
