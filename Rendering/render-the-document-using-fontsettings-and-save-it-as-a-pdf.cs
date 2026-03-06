using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Fonts;

class RenderDocumentWithFontSettings
{
    static void Main()
    {
        // Path to the source document (DOCX, etc.).
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Path where the resulting PDF will be saved.
        string outputPath = @"C:\Docs\OutputDocument.pdf";

        // Load the document.
        Document doc = new Document(inputPath);

        // Create a FontSettings object to control font resolution.
        FontSettings fontSettings = new FontSettings();

        // Example: add a folder that contains custom fonts.
        // Adjust the folder path to where your fonts are located.
        string customFontsFolder = @"C:\Docs\CustomFonts";
        FontSourceBase folderSource = new FolderFontSource(customFontsFolder, true);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

        // Assign the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Create PDF save options (you can customize further if needed).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Example: embed full fonts to avoid subsetting.
            EmbedFullFonts = true,

            // Example: use high‑quality rendering.
            UseHighQualityRendering = true
        };

        // Save the document as PDF using the specified options.
        doc.Save(outputPath, pdfOptions);
    }
}
