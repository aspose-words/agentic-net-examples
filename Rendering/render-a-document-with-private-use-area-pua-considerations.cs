using System;
using Aspose.Words;
using Aspose.Words.Saving;

class RenderPuaDocument
{
    static void Main()
    {
        // Path where the output file will be saved.
        string artifactsDir = @"C:\Output\";

        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some regular text.
        builder.Writeln("This is normal text.");

        // Insert characters from the Unicode Private Use Area (PUA).
        // Example: U+E001, U+E002, U+E003.
        builder.Writeln("\uE001\uE002\uE003");

        // Keep the original font metrics after any font substitution.
        // This helps preserve the appearance of PUA glyphs.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Rebuild the page layout so that the changes are taken into account
        // before rendering or saving to a fixed‑page format.
        doc.UpdatePageLayout();

        // Configure PDF save options.
        // Using DmlRenderingMode.Fallback ensures that any DrawingML shapes
        // are rendered via their fallback representations, which can be safer
        // for documents that contain custom glyphs.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DmlRenderingMode = DmlRenderingMode.Fallback
        };

        // Save the document to PDF.
        doc.Save(artifactsDir + "PUA_Rendered.pdf", pdfOptions);
    }
}
