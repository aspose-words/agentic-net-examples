using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsDocmConversion
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCM file.
            Document doc = new Document(@"C:\Input\source.docm");

            // -----------------------------------------------------------------
            // Save as DOCM (macro‑enabled) using OoxmlSaveOptions.
            // -----------------------------------------------------------------
            OoxmlSaveOptions docmSaveOptions = new OoxmlSaveOptions(SaveFormat.Docm)
            {
                // Example: keep the generator name in the output file.
                ExportGeneratorName = true,
                // Example: embed fonts if needed.
                AllowEmbeddingPostScriptFonts = false
            };
            doc.Save(@"C:\Output\result.docm", docmSaveOptions);

            // -----------------------------------------------------------------
            // Save as HTML with custom HtmlSaveOptions.
            // -----------------------------------------------------------------
            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGeneratorName = true,          // embed Aspose.Words generator info
                ExportImagesAsBase64 = true,         // embed images directly in the HTML
                CssClassNamePrefix = "aw",           // prefix for generated CSS classes
                PrettyFormat = true                  // make the HTML more readable
            };
            doc.Save(@"C:\Output\result.html", htmlSaveOptions);

            // -----------------------------------------------------------------
            // Save as MHTML (Web archive) using HtmlSaveOptions with Mhtml format.
            // -----------------------------------------------------------------
            HtmlSaveOptions mhtmlSaveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                ExportGeneratorName = true,
                ExportCidUrlsForMhtmlResources = true, // use CID URLs for embedded resources
                PrettyFormat = true
            };
            doc.Save(@"C:\Output\result.mhtml", mhtmlSaveOptions);

            // -----------------------------------------------------------------
            // Save as EPUB using HtmlSaveOptions with Epub format.
            // -----------------------------------------------------------------
            HtmlSaveOptions epubSaveOptions = new HtmlSaveOptions(SaveFormat.Epub)
            {
                ExportGeneratorName = true,
                ExportImagesAsBase64 = true,
                CssSavingCallback = null, // default CSS handling
                PrettyFormat = true
            };
            doc.Save(@"C:\Output\result.epub", epubSaveOptions);
        }
    }
}
