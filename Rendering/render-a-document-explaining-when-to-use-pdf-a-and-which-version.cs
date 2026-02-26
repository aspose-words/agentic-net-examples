using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Title.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
        builder.Writeln("When to Use PDF/A and Which Version to Choose");

        // Reset to normal style for body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;

        // Introductory explanation.
        builder.Writeln("PDF/A is an ISO‑standardized version of PDF designed for long‑term archiving of electronic documents.");
        builder.Writeln("It guarantees that the visual appearance of the document remains consistent and that all resources required for rendering are embedded.");

        // Decision guide.
        builder.Writeln("Select a PDF/A version based on your specific needs:");
        // Start a bullet list.
        builder.ListFormat.ApplyBulletDefault();

        builder.Writeln("PDF/A‑1a (ISO 19005‑1): preserves visual appearance **and** document structure (tagged). Ideal when searchable, reusable content is required.");
        builder.Writeln("PDF/A‑1b (ISO 19005‑1): preserves only visual appearance. Smallest file size, suitable for simple archival.");
        builder.Writeln("PDF/A‑2u (ISO 19005‑2): adds Unicode text extraction and supports newer features such as JPEG2000. Use when reliable text extraction is needed.");
        builder.Writeln("PDF/A‑3u (ISO 19005‑3): same as PDF/A‑2u but also permits embedding of arbitrary file attachments. Useful for bundling source data with the PDF.");
        builder.Writeln("PDF/A‑4 (ISO 19005‑4): the most recent standard, combines features of PDF/A‑2/‑3 and adds support for PDF 2.0 capabilities.");

        // End the list.
        builder.ListFormat.RemoveNumbers();

        // Save the document as a PDF/A‑2u compliant PDF.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA2u; // PDF/A‑2u compliance.
        doc.Save("PdfA_Explanation.pdf", saveOptions);
    }
}
