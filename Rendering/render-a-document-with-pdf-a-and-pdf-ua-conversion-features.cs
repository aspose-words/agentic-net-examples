using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceExample
{
    static void Main()
    {
        // Create a new document and add some sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as PDF/A-4 and PDF/UA-2 compliant.");

        // -------------------------------------------------
        // Save as PDF/A-4 (ISO 19005-4) – visual preservation.
        // -------------------------------------------------
        PdfSaveOptions pdfA4Options = new PdfSaveOptions();
        pdfA4Options.Compliance = PdfCompliance.PdfA4; // Set compliance level to PDF/A-4.
        // No need to force document structure for PDF/A-4, default settings are sufficient.
        doc.Save("Output_PdfA4.pdf", pdfA4Options);

        // ---------------------------------------------------------------
        // Save as PDF/A-4 + PDF/UA-2 (ISO 19005-4 + ISO 14289-2) – accessibility.
        // ---------------------------------------------------------------
        PdfSaveOptions pdfA4Ua2Options = new PdfSaveOptions();
        pdfA4Ua2Options.Compliance = PdfCompliance.PdfA4Ua2; // Set compliance to PDF/A-4 + PDF/UA-2.
        // PDF/UA requires document structure (tags) to be exported.
        pdfA4Ua2Options.ExportDocumentStructure = true;
        doc.Save("Output_PdfA4Ua2.pdf", pdfA4Ua2Options);
    }
}
