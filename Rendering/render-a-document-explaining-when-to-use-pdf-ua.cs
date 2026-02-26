using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Build the content that explains when PDF/UA should be used.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("When to use PDF/UA");
        builder.Writeln();
        builder.Writeln("PDF/UA (Universal Accessibility) is a standard that ensures PDF documents are accessible to people with disabilities.");
        builder.Writeln("Use PDF/UA when:");
        builder.Writeln("1. The document must comply with accessibility regulations (e.g., Section 508, EN 301 549).");
        builder.Writeln("2. The content needs to be searchable and navigable by assistive technologies.");
        builder.Writeln("3. You need to guarantee that the document structure (tags) is preserved.");
        builder.Writeln("4. The PDF will be distributed in environments where accessibility is a legal requirement.");
        builder.Writeln("5. You want to future‑proof the document for inclusive use.");

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Choose the PDF/UA compliance level (PDF/UA‑1 or PDF/UA‑2).
            Compliance = PdfCompliance.PdfUa1,
            // Required for PDF/UA: display the document title in the viewer's title bar.
            DisplayDocTitle = true,
            // Export document structure (tags); this is automatically required for PDF/UA.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF with the specified PDF/UA options.
        doc.Save("PdfUaExplanation.pdf", saveOptions);
    }
}
