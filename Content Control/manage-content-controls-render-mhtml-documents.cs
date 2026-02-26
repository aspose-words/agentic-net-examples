using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source document.
        Document doc = new Document("Rendering.docx");

        // Save the document as MHTML using CID URLs for resources.
        HtmlSaveOptions cidOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = true,   // Use "cid:" URLs.
            CssStyleSheetType = CssStyleSheetType.External,
            ExportFontResources = true,
            PrettyFormat = true
        };
        string cidOutputPath = "OutputWithCid.mht";
        doc.Save(cidOutputPath, cidOptions);

        // Save the document as MHTML using the default file‑name references.
        HtmlSaveOptions defaultOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            ExportCidUrlsForMhtmlResources = false,  // Use file names.
            CssStyleSheetType = CssStyleSheetType.External,
            ExportFontResources = true,
            PrettyFormat = true
        };
        string defaultOutputPath = "OutputWithoutCid.mht";
        doc.Save(defaultOutputPath, defaultOptions);
    }
}
