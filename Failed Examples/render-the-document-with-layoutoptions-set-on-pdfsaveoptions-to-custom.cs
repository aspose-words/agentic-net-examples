// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;
using Aspose.Words.Tables;

class RenderDocumentWithLayout
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add some content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("First paragraph with custom layout settings.");
        builder.Writeln("Second paragraph with the same settings.");

        // -----------------------------------------------------------------
        // Configure page layout: size, margins, orientation, columns.
        // -----------------------------------------------------------------
        // Access the first (and only) section.
        var pageSetup = doc.Sections[0].PageSetup;

        // Set custom page size (A4 in points).
        pageSetup.PageWidth = 595;   // 210 mm
        pageSetup.PageHeight = 842;  // 297 mm

        // Set orientation.
        pageSetup.Orientation = Orientation.Landscape;

        // Set custom margins (50 points on each side).
        pageSetup.LeftMargin = 50;
        pageSetup.RightMargin = 50;
        pageSetup.TopMargin = 50;
        pageSetup.BottomMargin = 50;

        // Set two columns with 20 points spacing between them.
        pageSetup.TextColumns.Clear();               // Ensure collection is empty.
        pageSetup.TextColumns.Add(2, 20);            // 2 columns, 20pt spacing.

        // -----------------------------------------------------------------
        // Configure paragraph formatting: line spacing and paragraph spacing.
        // -----------------------------------------------------------------
        // Apply the same formatting to all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.LineSpacing = 18;      // 1.5 line spacing (approx.).
            para.ParagraphFormat.SpaceBefore = 12;      // 12 points before.
            para.ParagraphFormat.SpaceAfter = 12;       // 12 points after.
        }

        // -----------------------------------------------------------------
        // Update the layout to apply the changes.
        // -----------------------------------------------------------------
        doc.UpdatePageLayout();

        // -----------------------------------------------------------------
        // Create PdfSaveOptions and save the document as PDF.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: display pages in a single column when opened in a PDF viewer.
            PageLayout = PdfPageLayout.OneColumn
        };

        // Save the document to a PDF file using the specified options.
        doc.Save("RenderedDocument.pdf", pdfOptions);
    }
}
