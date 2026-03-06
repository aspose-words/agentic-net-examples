// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Configure page layout -----
        // Access the first (and only) section.
        var section = doc.Sections[0];
        var pageSetup = section.PageSetup;

        // Set page size to A4 (in points: 1 inch = 72 points).
        pageSetup.PageWidth = 595;   // 8.27 inches
        pageSetup.PageHeight = 842;  // 11.69 inches

        // Set orientation to Landscape.
        pageSetup.Orientation = Orientation.Landscape;

        // Set custom margins (1 inch on each side).
        pageSetup.Margins.Top = 72;
        pageSetup.Margins.Bottom = 72;
        pageSetup.Margins.Left = 72;
        pageSetup.Margins.Right = 72;

        // Set two text columns.
        pageSetup.TextColumns.SetCount(2);

        // ----- Configure paragraph formatting -----
        // Line spacing (1.5 lines ≈ 18 points) and paragraph spacing.
        builder.ParagraphFormat.LineSpacing = 18;      // 1.5 lines
        builder.ParagraphFormat.SpaceBefore = 6;      // 6 points before each paragraph
        builder.ParagraphFormat.SpaceAfter = 12;      // 12 points after each paragraph

        // Add sample content.
        builder.Writeln("First paragraph with custom line spacing and paragraph spacing.");
        builder.Writeln("Second paragraph follows the same formatting.");

        // Rebuild the layout after making changes.
        doc.UpdatePageLayout();

        // ----- Create PDF save options -----
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use high‑quality rendering for better visual fidelity.
            UseHighQualityRendering = true,

            // Set how the PDF viewer should display pages (single column view).
            PageLayout = PdfPageLayout.OneColumn
        };

        // Save the document as PDF using the configured options.
        doc.Save("Output.pdf", pdfOptions);
    }
}
