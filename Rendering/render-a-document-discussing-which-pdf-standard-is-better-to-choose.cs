using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfStandardComparison
{
    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Use DocumentBuilder to add content that discusses PDF standards.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Title
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Title;
            builder.Writeln("Choosing the Right PDF Standard");

            // Introduction
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Introduction");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PDF (Portable Document Format) has several compliance levels, each designed for different use cases. " +
                "Selecting the appropriate standard ensures long‑term preservation, accessibility, and compatibility.");

            // PDF 1.7
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF 1.7 (ISO 32000‑1)");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "The baseline PDF 1.7 standard is suitable for general document exchange. " +
                "It does not enforce any archival or accessibility requirements, making it the most flexible choice for everyday use.");

            // PDF/A family
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF/A – Archival Standards");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PDF/A is intended for long‑term preservation. " +
                "PDF/A‑1a and PDF/A‑2a include document structure (tagging) for searchable content, while PDF/A‑1b and PDF/A‑2u focus on visual fidelity. " +
                "If you need guaranteed rendering over decades, choose a PDF/A level that matches your archival policy.");

            // PDF/UA – Accessibility
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF/UA – Accessibility");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PDF/UA (Universal Accessibility) ensures that PDFs are usable by assistive technologies. " +
                "When accessibility is a requirement, combine PDF/UA with an archival level such as PDF/A‑4f or PDF/A‑4Ua2.");

            // Recommendation
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Recommendation");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "• For general distribution: **PDF 1.7** (PdfCompliance.Pdf17). " +
                "• For archival without accessibility: **PDF/A‑2u** (PdfCompliance.PdfA2u). " +
                "• For archival with full accessibility: **PDF/A‑4 + PDF/UA‑2** (PdfCompliance.PdfA4Ua2).");

            // Create PDF save options and set the desired compliance level.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // Change this value to test different standards.
                // Example: PdfCompliance.PdfA2u for archival, PdfCompliance.PdfA4Ua2 for archival + accessibility.
                Compliance = PdfCompliance.PdfA4Ua2
            };

            // Save the document as a PDF using the specified compliance.
            doc.Save("PdfStandardComparison.pdf", saveOptions);
        }
    }
}
