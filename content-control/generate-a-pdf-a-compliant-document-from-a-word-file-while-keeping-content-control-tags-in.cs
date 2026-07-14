using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

namespace ContentControlPdfAExample
{
    public class Program
    {
        public static void Main()
        {
            // Define file names in the working directory.
            const string docxPath = "sample.docx";
            const string pdfPath = "sample-pdfa.pdf";

            // -----------------------------------------------------------------
            // 1. Create a Word document with a plain‑text content control.
            // -----------------------------------------------------------------
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph before the content control.
            builder.Writeln("This document contains a content control that will be kept in the PDF/A output.");

            // Create an inline plain‑text StructuredDocumentTag (content control).
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",
                Tag = "customer-name"
            };
            // Set placeholder text (optional) and initial content.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "John Doe"));

            // Insert the content control into the current paragraph.
            builder.InsertNode(sdt);
            builder.Writeln(); // Move to a new line after the control.

            // Save the seed DOCX file.
            doc.Save(docxPath);

            // -----------------------------------------------------------------
            // 2. Load the DOCX file and convert it to PDF/A.
            // -----------------------------------------------------------------
            Document loadedDoc = new Document(docxPath);

            // Configure PDF save options for PDF/A‑1a compliance.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a,
                // Preserve form fields so that the content control appears as an interactive field in the PDF.
                PreserveFormFields = true,
                // Use the content control's Tag as the form field name (helps keep the mapping clear).
                UseSdtTagAsFormFieldName = true,
                // Export document structure is required for PDF/A‑1a; the property is ignored but set for clarity.
                ExportDocumentStructure = true
            };

            // Save the document as PDF/A.
            loadedDoc.Save(pdfPath, pdfOptions);

            // Inform the user (no interactive prompts, just console output).
            Console.WriteLine($"DOCX file saved to: {Path.GetFullPath(docxPath)}");
            Console.WriteLine($"PDF/A file saved to: {Path.GetFullPath(pdfPath)}");
        }
    }
}
