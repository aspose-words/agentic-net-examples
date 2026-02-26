using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Saving;

class StructuredDocumentTagExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text Structured Document Tag (content control) at the current cursor position.
        // The markup level (inline) is detected automatically based on the insertion point.
        StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);

        // Set placeholder text to demonstrate the tag's appearance.
        sdt.Title = "SampleTag";
        sdt.Tag = "SampleTag";
        sdt.IsShowingPlaceholderText = true;

        // Write some content inside the tag.
        builder.Writeln("Content inside the Structured Document Tag.");

        // Prepare PDF save options to export the document structure (tags) to the PDF.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true // Enables PDF tags for accessibility.
        };

        // Save the document as PDF with the specified options.
        doc.Save("StructuredDocumentTag_Output.pdf", pdfOptions);
    }
}
