using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source EPUB document.
        Document doc = new Document("input.epub");

        // Configure save options for MHTML.
        // - Use the MHTML format.
        // - Embed images as Base64 so they are stored inside the MHTML file.
        // - Export font resources and embed them.
        // - Use CID URLs for resources (helps some mail agents/browsers).
        // - Keep CSS inline and format the output nicely.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportImagesAsBase64 = true,
            ExportFontResources = true,
            ExportCidUrlsForMhtmlResources = true,
            CssStyleSheetType = CssStyleSheetType.Inline,
            PrettyFormat = true,
            Encoding = Encoding.UTF8
        };

        // Save the document as MHTML with all resources embedded.
        doc.Save("output.mht", saveOptions);
    }
}
