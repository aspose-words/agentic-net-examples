using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;
using Aspose.Words.Loading; // Added for LoadOptions

class Program
{
    static void Main()
    {
        // Folder that contains the custom fonts you want to use.
        string fontsFolder = @"C:\MyFonts";

        // Path to the source Word document.
        string sourceDocPath = @"C:\Docs\Sample.docx";

        // Path where the resulting PDF will be saved.
        string pdfPath = @"C:\Docs\Sample.pdf";

        // -----------------------------------------------------------------
        // 1. Create FontSettings and point it to the custom fonts folder.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        // true => search subfolders as well.
        fontSettings.SetFontsFolder(fontsFolder, true);

        // -----------------------------------------------------------------
        // 2. Load the document, applying the FontSettings via LoadOptions.
        // -----------------------------------------------------------------
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings; // Apply font settings during load.

        Document doc = new Document(sourceDocPath, loadOptions);

        // If you create a blank document instead, assign the FontSettings directly:
        // Document doc = new Document();
        // doc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 3. Prepare PDF save options (optional customizations).
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example: embed the full fonts into the PDF.
        pdfOptions.EmbedFullFonts = true;

        // -----------------------------------------------------------------
        // 4. Save the document as PDF using the configured options.
        // -----------------------------------------------------------------
        doc.Save(pdfPath, pdfOptions);
    }
}
