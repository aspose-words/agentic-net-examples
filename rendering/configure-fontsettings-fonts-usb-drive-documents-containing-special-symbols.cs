using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontSettingsFromUsb
{
    static void Main()
    {
        // Path to the folder on the USB drive that contains the required TrueType fonts.
        // Adjust the drive letter and folder name as needed.
        string usbFontsFolder = @"E:\MyFonts";

        // Create a FontSettings instance and configure it to search the USB folder.
        // The second argument (true) enables recursive scanning of subfolders.
        FontSettings usbFontSettings = new FontSettings();
        usbFontSettings.SetFontsFolder(usbFontsFolder, true);

        // Apply the FontSettings to LoadOptions so that fonts are resolved while loading.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = usbFontSettings;

        // Load the source document using the configured LoadOptions.
        // Replace "input.docx" with the actual path to your document.
        Document doc = new Document("input.docx", loadOptions);

        // Optional: add some text that uses special symbols to verify font loading.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "CustomFontFromUsb"; // use a font that exists in the USB folder
        builder.Writeln("Special symbols: Ω, 漢字, 😊");

        // Save the processed document. The fonts will be resolved from the USB drive.
        // Replace "output.pdf" with the desired output path and format.
        doc.Save("output.pdf");
    }
}
