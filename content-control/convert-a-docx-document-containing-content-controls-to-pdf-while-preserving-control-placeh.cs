using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a description paragraph.
        builder.Writeln("Below is a plain‑text content control with a placeholder:");

        // Create an inline plain‑text content control (SDT).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name",
            IsShowingPlaceholderText = true
        };

        // Define the placeholder text that will be shown when the control is empty.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Enter name here"));

        // Insert the content control into the document.
        builder.InsertNode(sdt);
        builder.Writeln(); // End the paragraph.

        // Save the source DOCX file.
        const string docxPath = "sample.docx";
        doc.Save(docxPath);

        // Load the DOCX file (simulating an existing document with content controls).
        Document loadedDoc = new Document(docxPath);

        // Configure PDF save options to preserve form fields and use the SDT tag as the field name.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PreserveFormFields = true,
            UseSdtTagAsFormFieldName = true
        };

        // Convert the document to PDF while preserving the content control placeholder.
        const string pdfPath = "output.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
