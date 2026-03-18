using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path where the resulting MHTML file will be saved.
        string outputPath = @"C:\Docs\Sample.mht";

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure save options for MHTML.
        // - CssStyleSheetType.External tells Aspose.Words to treat linked CSS as external files,
        //   which are then packaged inside the MHTML container.
        // - ExportCidUrlsForMhtmlResources forces resources (CSS, images, fonts) to be referenced
        //   via CID URLs, ensuring they are embedded in the MHTML output.
        // - ExportImagesAsBase64 and ExportFontsAsBase64 are optional but often desirable for a
        //   fully self‑contained MHTML file.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            CssStyleSheetType = CssStyleSheetType.External,
            ExportCidUrlsForMhtmlResources = true,
            ExportImagesAsBase64 = true,
            ExportFontsAsBase64 = true,
            PrettyFormat = true
        };

        // Save the document as MHTML; all linked CSS files are automatically embedded.
        doc.Save(outputPath, saveOptions);
    }
}
