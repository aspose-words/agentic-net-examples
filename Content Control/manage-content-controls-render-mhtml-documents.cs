using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load a DOCX file that contains content controls.
        // The Document constructor handles the loading; no custom load options are required here.
        Document doc = new Document("InputWithContentControls.docx");

        // Configure save options for MHTML output.
        // HtmlSaveOptions is used for both HTML and MHTML formats.
        HtmlSaveOptions mhtmlOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Preserve headers/footers per section so they appear correctly in the MHTML.
            ExportHeadersFootersMode = ExportHeadersFootersMode.PerSection,

            // Include built‑in and custom document properties in the output.
            ExportDocumentProperties = true,

            // Export font resources so that the document renders with the original fonts.
            ExportFontResources = true,

            // Use CID URLs for embedded resources (images, fonts, CSS) – required for proper MHTML packaging.
            ExportCidUrlsForMhtmlResources = true,

            // Keep images as external files rather than Base64 to reduce MHTML size.
            ExportImagesAsBase64 = false,

            // Save as XHTML (transitional) to ensure broad compatibility.
            HtmlVersion = HtmlVersion.Xhtml,
            ExportXhtmlTransitional = true,

            // Produce nicely indented output for easier debugging.
            PrettyFormat = true
        };

        // Save the document as an MHTML file using the configured options.
        doc.Save("OutputDocument.mht", mhtmlOptions);
    }
}
