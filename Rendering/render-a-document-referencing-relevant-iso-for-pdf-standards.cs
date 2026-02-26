using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Insert a paragraph that mentions the ISO standards the PDF will comply with.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF complies with ISO 32000-2 (PDF 2.0), ISO 19005-4 (PDF/A-4) and ISO 14289-2 (PDF/UA-2).");

        // Configure PDF save options.
        PdfSaveOptions saveOptions = new PdfSaveOptions();

        // Set the compliance level to PDF/A-4 + PDF/UA-2.
        // PDF/A-4 corresponds to ISO 19005-4 and PDF/UA-2 corresponds to ISO 14289-2.
        saveOptions.Compliance = PdfCompliance.PdfA4Ua2;

        // Export the document structure (tags) – required for PDF/UA compliance.
        saveOptions.ExportDocumentStructure = true;

        // Render DrawingML shapes directly (no fallback substitution).
        saveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

        // Save the document as a PDF using the configured options.
        doc.Save("CompliantDocument.pdf", saveOptions);
    }
}
