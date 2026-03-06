using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfUaExample
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document title – required for PDF/UA viewers.
        doc.BuiltInDocumentProperties.Title = "When to Use PDF/UA";

        // Add a heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("When to Use PDF/UA");

        // Add explanatory paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("PDF/UA (Universal Accessibility) is a PDF standard that ensures the document is accessible to people with disabilities. Use PDF/UA when:");
        builder.Writeln("- The document must comply with accessibility regulations (e.g., Section 508, EN 301 549).");
        builder.Writeln("- You need to guarantee that assistive technologies can read the content, navigate headings, tables, and form fields.");
        builder.Writeln("- The document will be distributed publicly or within an organization that requires accessible PDFs.");
        builder.Writeln("- You want searchable, tagged content that can be repurposed.");

        // Add a list of typical scenarios.
        builder.Writeln("Typical scenarios:");
        builder.Writeln("1. Government forms and reports.");
        builder.Writeln("2. Educational material for students with disabilities.");
        builder.Writeln("3. Corporate policies that must meet legal accessibility standards.");
        builder.Writeln("4. E‑books and manuals intended for a broad audience.");

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Choose PDF/UA-1 compliance (use PdfUa2 for the newer standard).
            Compliance = PdfCompliance.PdfUa1,
            // Required for PDF/UA: display the document title in the viewer's title bar.
            DisplayDocTitle = true,
            // Export document structure (tags); automatically enabled for PDF/UA.
            ExportDocumentStructure = true
        };

        // Save the document as a PDF that complies with PDF/UA.
        doc.Save("PdfUaExample.pdf", saveOptions);
    }
}
