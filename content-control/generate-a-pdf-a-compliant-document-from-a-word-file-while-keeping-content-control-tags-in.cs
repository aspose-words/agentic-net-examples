using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the working directory.
        const string wordFile = "sample.docx";
        const string pdfFile = "sample-pdfa.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create a Word document with a plain‑text content control.
        // -----------------------------------------------------------------
        Document doc = new Document();
        // The first paragraph is created automatically in a new document.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create an inline plain‑text StructuredDocumentTag (content control).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        // Set the initial text inside the content control.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso Ltd."));

        // Insert the content control into the paragraph.
        paragraph.AppendChild(sdt);

        // Save the Word document to disk.
        doc.Save(wordFile);

        // -----------------------------------------------------------------
        // Step 2: Load the Word document and convert it to PDF/A.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(wordFile);

        // Configure PDF save options for PDF/A‑1a compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑1a includes visual fidelity and document structure.
            Compliance = PdfCompliance.PdfA1a,
            // Preserve content controls as interactive form fields in the PDF.
            PreserveFormFields = true,
            // Use the Tag property of the content control as the form field name.
            UseSdtTagAsFormFieldName = true,
            // Export the document structure (required for PDF/A‑1a, but set explicitly).
            ExportDocumentStructure = true
        };

        // Save the document as a PDF/A compliant file.
        loadedDoc.Save(pdfFile, pdfOptions);

        // -----------------------------------------------------------------
        // Optional: Inform the user via console (no input required).
        // -----------------------------------------------------------------
        Console.WriteLine($"Word file created: {Path.GetFullPath(wordFile)}");
        Console.WriteLine($"PDF/A file created: {Path.GetFullPath(pdfFile)}");
    }
}
