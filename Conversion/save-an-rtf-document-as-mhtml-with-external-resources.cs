using System;
using Aspose.Words;
using Aspose.Words.Saving;

class SaveRtfAsMhtml
{
    static void Main()
    {
        // Path to the source RTF document.
        string rtfPath = @"C:\Docs\input.rtf";

        // Path where the MHTML file will be saved.
        string mhtmlPath = @"C:\Docs\output.mht";

        // Load the RTF document.
        Document doc = new Document(rtfPath);

        // Configure save options for MHTML with external resources.
        HtmlSaveOptions options = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export CSS to an external stylesheet instead of embedding it.
            CssStyleSheetType = CssStyleSheetType.External,

            // Export font files as separate resources.
            ExportFontResources = true,

            // Keep images as separate files (not Base64‑encoded).
            ExportImagesAsBase64 = false,

            // Use file‑name URLs for resources (default behavior).
            ExportCidUrlsForMhtmlResources = false,

            // Optional: make the output more readable.
            PrettyFormat = true
        };

        // Save the document as MHTML using the configured options.
        doc.Save(mhtmlPath, options);
    }
}
