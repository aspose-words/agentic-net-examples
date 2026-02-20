using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOC document.
        Document doc = new Document("input.doc");

        // Configure save options to produce MHTML output.
        // The SaveFormat.Mhtml enum value tells Aspose.Words to save in MHTML (Web archive) format.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Use HTML5 standard for the generated markup.
            HtmlVersion = HtmlVersion.Html5,
            // Export images as embedded Base64 data.
            ExportImagesAsBase64 = true,
            // Embed CSS directly into the MHTML (inline stylesheet).
            CssStyleSheetType = CssStyleSheetType.Inline
        };

        // Save the document as an MHTML file.
        doc.Save("output.mht", saveOptions);
    }
}
