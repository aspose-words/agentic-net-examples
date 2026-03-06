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
                "PDF files can be saved in several compliance levels, each serving different purposes. " +
                "Selecting the appropriate standard depends on factors such as long‑term preservation, " +
                "accessibility, and compatibility with existing workflows.");

            // PDF 1.7 vs PDF 2.0
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF 1.7 vs PDF 2.0");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "Pdf17 (ISO 32000‑1) is the original PDF 1.7 specification. " +
                "Pdf20 (ISO 32000‑2) adds new features, improved security, and better handling of color spaces. " +
                "If you need the latest PDF capabilities and your downstream tools support PDF 2.0, choose Pdf20; " +
                "otherwise Pdf17 remains the safest choice for maximum compatibility.");

            // PDF/A family
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF/A – Archival Standards");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PDF/A is designed for long‑term preservation. " +
                "PdfA1b ensures visual fidelity, while PdfA1a adds document structure for searchability. " +
                "PdfA2u and PdfA3u extend these capabilities with Unicode text extraction and, for PdfA3u, embedded attachments. " +
                "PdfA4 and PdfA4f are the latest, supporting modern features and optional attachments. " +
                "Choose a PDF/A level when you must guarantee that the document can be rendered exactly the same way years from now.");

            // PDF/UA – Accessibility
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("PDF/UA – Accessibility");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PdfUa1 and PdfUa2 enforce accessibility requirements, making PDFs usable with assistive technologies. " +
                "If your audience includes users with disabilities or you must comply with accessibility regulations, " +
                "select a PDF/UA level (PdfUa2 is the newer standard).");

            // Combined standards
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
            builder.Writeln("Combined Standards");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "PdfA4Ua2 combines PDF/A‑4 archival features with PDF/UA‑2 accessibility, offering the best of both worlds. " +
                "Use this when you need both long‑term preservation and full accessibility.");

            // Recommendation summary
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("Recommendation Summary");
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln(
                "- For general purpose documents with broad compatibility: **Pdf17**. " +
                "If you need newer features and your tools support it: **Pdf20**.\n" +
                "- For archival purposes: **PdfA2u** (or **PdfA4** for the latest features). " +
                "Add **PdfA4f** if you need to embed attachments.\n" +
                "- For accessibility compliance: **PdfUa2** (or **PdfA4Ua2** to combine archiving and accessibility).");

            // Create PDF save options and set the desired compliance level.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // Change this value to test different standards.
                // Example: PdfCompliance.PdfA2u for archival PDF/A‑2u compliance.
                Compliance = PdfCompliance.PdfA2u
            };

            // Save the document as a PDF using the specified compliance level.
            doc.Save("PdfStandardComparison.pdf", saveOptions);
        }
    }
}
