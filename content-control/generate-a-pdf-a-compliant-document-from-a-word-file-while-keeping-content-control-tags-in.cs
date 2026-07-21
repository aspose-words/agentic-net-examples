using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Define file paths in the current working directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.docx");
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "sample-pdfa.pdf");

        // -----------------------------------------------------------------
        // 1. Create a Word document that contains several content controls.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Plain‑text content control (inline).
        StructuredDocumentTag plainTextSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        plainTextSdt.RemoveAllChildren();
        plainTextSdt.AppendChild(new Run(doc, "Contoso"));

        // Move to a new paragraph and insert the inline SDT.
        builder.Writeln();
        builder.InsertNode(plainTextSdt);
        builder.Writeln(); // New line after the inline SDT.

        // Rich‑text (block‑level) content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "Comments",
            Tag = "comments"
        };
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Enter your comments here..."));
        richTextSdt.AppendChild(richParagraph);

        // Append the block‑level SDT directly to the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);
        builder.Writeln(); // New line after the block‑level SDT.

        // Checkbox content control (inline).
        StructuredDocumentTag checkBoxSdt = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "AgreeTerms",
            Tag = "agree-terms",
            Checked = false
        };
        builder.Writeln();
        builder.InsertNode(checkBoxSdt);
        builder.Write(" I agree to the terms and conditions.");
        builder.Writeln();

        // Save the seed document.
        doc.Save(inputPath);

        // -----------------------------------------------------------------
        // 2. Load the document and convert it to PDF/A‑1a while preserving
        //    the content controls as interactive form fields.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputPath);

        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // PDF/A‑1a compliance (includes PDF/A‑1b + document structure).
            Compliance = PdfCompliance.PdfA1a,

            // Preserve Word content controls as PDF form fields.
            PreserveFormFields = true,

            // Use the Tag property of the SDT as the name of the PDF form field.
            UseSdtTagAsFormFieldName = true,

            // Export document structure (required for PDF/A‑1a, but ignored automatically).
            ExportDocumentStructure = true
        };

        // Save the PDF/A document.
        loadedDoc.Save(outputPdfPath, pdfOptions);
    }
}
