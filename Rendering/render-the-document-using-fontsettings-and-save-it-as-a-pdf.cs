using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source document (replace with actual file).
        string inputPath = "Input.docx";

        // Path where the PDF will be saved.
        string outputPath = "Output.pdf";

        // Folder that contains custom fonts to be used during rendering.
        string fontsFolder = "Fonts";

        // Load the document using the Document constructor (lifecycle rule).
        Document doc = new Document(inputPath);

        // Create FontSettings and add a folder source for custom fonts.
        FontSettings fontSettings = new FontSettings();

        // Preserve the default system font source and add our custom folder source.
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();
        FontSourceBase folderSource = new FolderFontSource(fontsFolder, true);
        fontSettings.SetFontsSources(new FontSourceBase[] { originalSources[0], folderSource });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Create PdfSaveOptions to control PDF rendering.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Example option: embed the full fonts into the PDF.
        pdfOptions.EmbedFullFonts = true;

        // Example option: use high‑quality rendering (slower but better visual quality).
        pdfOptions.UseHighQualityRendering = true;

        // Save the document as PDF using the Save method with options (lifecycle rule).
        doc.Save(outputPath, pdfOptions);
    }
}
