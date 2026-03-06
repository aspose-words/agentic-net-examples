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
        // Paths to the input document, output PDF and the folder containing TrueType fonts.
        string inputDocPath = @"C:\Docs\input.docx";
        string outputPdfPath = @"C:\Docs\output.pdf";
        string trueTypeFontsFolder = @"C:\Fonts";

        // Load the source document.
        Document doc = new Document(inputDocPath);

        // Preserve the original font sources.
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add a folder font source that points to the TrueType fonts location.
        FolderFontSource folderSource = new FolderFontSource(trueTypeFontsFolder, true);
        FontSettings.DefaultInstance.SetFontsSources(originalSources.Concat(new[] { folderSource }).ToArray());

        // Configure PDF save options (e.g., embed full fonts).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EmbedFullFonts = true // Embed all glyphs of the used fonts.
        };

        // Save the document as PDF using the configured options.
        doc.Save(outputPdfPath, pdfOptions);

        // (Optional) Restore the original font sources if further processing is needed.
        FontSettings.DefaultInstance.SetFontsSources(originalSources);
    }
}
