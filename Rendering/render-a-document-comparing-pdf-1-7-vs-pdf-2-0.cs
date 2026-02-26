using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfComplianceComparison
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First section – PDF 1.7 description.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("PDF 1.7 (ISO 32000‑1)");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("PDF 1.7 is the original PDF specification defined by ISO 32000‑1. " +
                        "It is widely supported by PDF viewers and libraries. " +
                        "Features include basic text, images, annotations, and standard security.");

        // Add a page break to separate the two sections.
        builder.InsertBreak(BreakType.PageBreak);

        // Second section – PDF 2.0 description.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("PDF 2.0 (ISO 32000‑2)");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("PDF 2.0 is the newer PDF specification defined by ISO 32000‑2. " +
                        "It adds support for richer color spaces, enhanced encryption, " +
                        "metadata improvements, and better accessibility features.");

        // -----------------------------------------------------------------
        // Save the document as PDF 1.7 compliant.
        PdfSaveOptions pdf17Options = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf17
        };
        doc.Save("PdfComparison_Pdf17.pdf", pdf17Options);

        // Save the same document as PDF 2.0 compliant.
        PdfSaveOptions pdf20Options = new PdfSaveOptions
        {
            Compliance = PdfCompliance.Pdf20
        };
        doc.Save("PdfComparison_Pdf20.pdf", pdf20Options);
    }
}
