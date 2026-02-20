using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Layout;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("input.docx");

        // Preserve original font metrics so that Private Use Area (PUA) characters retain their intended appearance.
        // The LayoutOptions property is read‑only, but its members are mutable.
        doc.LayoutOptions.KeepOriginalFontMetrics = true;

        // Configure save options to update ambiguous text fonts, which helps correctly render PUA characters.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            UpdateAmbiguousTextFont = true
        };

        // Save the document to PDF using the configured options.
        doc.Save("output.pdf", saveOptions);
    }
}
