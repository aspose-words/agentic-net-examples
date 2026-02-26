using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class RenderDocumentWithCustomFonts
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"C:\Docs\input.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\output.pdf";

        // Folder that contains the TrueType fonts to be used during rendering.
        string trueTypeFontsFolder = @"C:\MyFonts";

        // Load the document.
        Document doc = new Document(inputPath);

        // Preserve the original font sources so they can be restored later.
        FontSourceBase[] originalFontSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add a folder font source that points to the TrueType fonts location.
        FolderFontSource folderSource = new FolderFontSource(trueTypeFontsFolder, true);
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources.Concat(new[] { folderSource }).ToArray());

        // Configure PDF save options (e.g., embed full fonts).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true   // Embed complete fonts to ensure correct rendering.
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPath, pdfOptions);

        // Restore the original font sources to avoid side‑effects on other operations.
        FontSettings.DefaultInstance.SetFontsSources(originalFontSources);
    }
}
