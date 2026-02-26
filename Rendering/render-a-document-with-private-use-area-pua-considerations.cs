using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderDocumentWithPua
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a line that contains a Unicode Private Use Area character (U+E000).
        // This character is often used for custom glyphs that are not part of the standard Unicode set.
        string privateUseChar = "\uE000";
        builder.Writeln($"This line contains a Private Use Area character: {privateUseChar}");

        // Ensure that the original font metrics are kept after any font substitution.
        // This helps preserve the appearance of private use glyphs.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Rebuild the page layout so that the changes above are taken into account
        // before rendering or saving the document.
        doc.UpdatePageLayout();

        // Configure PDF save options.
        // DmlRenderingMode.DrawingML renders DrawingML shapes directly (default).
        // Keeping the default is sufficient for most cases, but the option is shown here explicitly.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.DrawingML,
            // Optional: embed the full fonts to ensure the private glyph is available in the PDF.
            EmbedFullFonts = true
        };

        // Save the document to PDF. The private use character will be rendered using the
        // original font metrics and embedded font, preserving its visual representation.
        doc.Save("RenderedWithPua.pdf", pdfOptions);
    }
}
