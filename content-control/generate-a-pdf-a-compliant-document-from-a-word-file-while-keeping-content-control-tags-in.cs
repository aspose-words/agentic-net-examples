using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a sample DOCX with a plain‑text content control (SDT).
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a heading.
        builder.Writeln("Demo document with a content control");

        // Create an inline plain‑text StructuredDocumentTag.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "SampleControl",
            Tag = "SampleTag",
            // Optional: show placeholder text when empty.
            IsShowingPlaceholderText = true
        };

        // Insert the SDT into the current paragraph.
        builder.InsertNode(sdt);

        // Add placeholder text inside the SDT.
        sdt.AppendChild(new Run(doc, "Enter text here..."));

        // Save the DOCX to the local file system.
        const string docPath = "Sample.docx";
        doc.Save(docPath);

        // Step 2: Load the DOCX and convert it to a PDF/A‑1a compliant PDF,
        // preserving the content control as a form field.
        Document loadedDoc = new Document(docPath);

        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑1a compliance (includes PDF/A‑1b + document structure).
            Compliance = PdfCompliance.PdfA1a,
            // Preserve Word form fields (content controls) as PDF form fields.
            PreserveFormFields = true,
            // Use the SDT Tag as the name of the PDF form field.
            UseSdtTagAsFormFieldName = true
        };

        const string pdfPath = "Sample_PdfA.pdf";
        loadedDoc.Save(pdfPath, pdfOptions);
    }
}
