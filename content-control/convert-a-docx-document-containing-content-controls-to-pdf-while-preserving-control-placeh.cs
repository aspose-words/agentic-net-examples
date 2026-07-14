using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a DOCX document with a plain‑text content control that has a placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some introductory text.
        builder.Writeln("Sample document with a content control placeholder:");

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name",
            IsShowingPlaceholderText = true // Show placeholder when the control is empty.
        };

        // Define the placeholder text.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Enter name here"));

        // Insert the content control into the first paragraph.
        Paragraph firstParagraph = doc.FirstSection.Body.FirstParagraph;
        firstParagraph.AppendChild(sdt);

        // Save the source DOCX file.
        const string docxPath = "input.docx";
        doc.Save(docxPath);

        // Load the DOCX file that contains the content control.
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options to preserve form fields (content controls) as interactive fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,
            UseSdtTagAsFormFieldName = true // Use the Tag as the PDF form field name.
        };

        // Convert and save the document to PDF while keeping the placeholder visible.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
