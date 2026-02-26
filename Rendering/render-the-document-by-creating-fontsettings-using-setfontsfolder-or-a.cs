using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading; // Added for LoadOptions
using Aspose.Words.Saving;

class FontSettingsPdfExample
{
    static void Main()
    {
        // Path to the source document.
        string inputPath = @"C:\Docs\Sample.docx";

        // Path to the folder that contains the custom fonts.
        // Set the second argument to true to search subfolders as well.
        string fontsFolder = @"C:\CustomFonts";

        // Create a FontSettings instance and point it to the custom fonts folder.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolder(fontsFolder, recursive: true);

        // Load the document using LoadOptions that carry the FontSettings.
        // This ensures that the fonts are available during layout and rendering.
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings
        };
        Document doc = new Document(inputPath, loadOptions);

        // Alternatively, you could assign the FontSettings after loading:
        // doc.FontSettings = fontSettings;

        // Create PdfSaveOptions – this object can be used to control PDF output.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Path to the output PDF file.
        string outputPath = @"C:\Docs\Sample.pdf";

        // Save the document as PDF using the options defined above.
        doc.Save(outputPath, pdfOptions);
    }
}
