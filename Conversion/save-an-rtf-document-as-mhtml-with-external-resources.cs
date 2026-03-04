using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source RTF document.
        string sourceRtfPath = @"C:\Docs\SourceDocument.rtf";

        // Load the RTF document.
        Document doc = new Document(sourceRtfPath);

        // Folder where external resources (images, CSS, fonts) will be written.
        string resourcesFolder = @"C:\Docs\ExternalResources";

        // Ensure the resources folder exists.
        Directory.CreateDirectory(resourcesFolder);

        // Configure save options for MHTML with external resources.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
        {
            // Export fonts as separate files.
            ExportFontResources = true,

            // Export CSS as an external stylesheet.
            CssStyleSheetType = CssStyleSheetType.External,

            // Do not embed images; they will be saved as separate files.
            ExportImagesAsBase64 = false,

            // Folder where images, CSS and fonts will be saved.
            ImagesFolder = resourcesFolder,

            // Use file name references (not CID URLs) for resources.
            ExportCidUrlsForMhtmlResources = false,

            // Optional: make the output HTML pretty-formatted.
            PrettyFormat = true
        };

        // Path for the resulting MHTML file.
        string outputMhtmlPath = @"C:\Docs\ResultDocument.mht";

        // Save the document as MHTML using the configured options.
        doc.Save(outputMhtmlPath, saveOptions);
    }
}
