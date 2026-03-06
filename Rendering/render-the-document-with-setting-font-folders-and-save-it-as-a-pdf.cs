using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the fonts folder, input document and output PDF.
        string fontsFolder = @"C:\MyFonts";
        string inputFile = @"C:\Docs\input.docx";
        string outputFile = @"C:\Docs\output.pdf";

        // Preserve the original font sources so we can restore them later.
        FontSourceBase[] originalSources = FontSettings.DefaultInstance.GetFontsSources();

        // Add a folder font source that points to the custom fonts folder.
        FolderFontSource customSource = new FolderFontSource(fontsFolder, true);
        FontSettings.DefaultInstance.SetFontsSources(new[] { originalSources[0], customSource });

        // Load the document to be rendered.
        Document doc = new Document(inputFile);

        // Configure PDF save options (high‑quality rendering and full font embedding).
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UseHighQualityRendering = true,
            EmbedFullFonts = true
        };

        // Save the document as a PDF using the configured options.
        doc.Save(outputFile, pdfOptions);

        // Restore the original font sources (optional cleanup).
        FontSettings.DefaultInstance.SetFontsSources(originalSources);
    }
}
