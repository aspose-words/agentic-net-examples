using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with a plain‑text content control that has a placeholder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some introductory text.
        builder.Writeln("Please fill in the following field:");

        // Insert an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name",
            // Show placeholder text when the control is empty.
            IsShowingPlaceholderText = true,
            // Lock the control so it becomes a PDF form field.
            LockContents = false
        };

        // The control is left empty; Word will display the placeholder.
        // Insert the SDT into the current paragraph.
        builder.InsertNode(sdt);

        // Add a line break after the control.
        builder.Writeln();

        // Save the DOCX to a local file.
        const string docxPath = "input.docx";
        doc.Save(docxPath);

        // Load the DOCX document.
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options to preserve form fields (content controls) as PDF form fields.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,
            // Use the Tag property of the SDT as the name of the PDF form field.
            UseSdtTagAsFormFieldName = true
        };

        // Save the document as PDF. The placeholder text will be visible in the PDF form field.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
