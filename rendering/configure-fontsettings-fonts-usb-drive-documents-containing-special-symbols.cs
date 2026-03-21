using System;
using System.IO;
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

        // Create a FontSettings instance and configure it to search the USB folder
        // only if the folder actually exists.
        FontSettings usbFontSettings = new FontSettings();
        if (Directory.Exists(usbFontsFolder))
        {
            // The second argument (true) enables recursive scanning of subfolders.
            usbFontSettings.SetFontsFolder(usbFontsFolder, true);
        }

        // Apply the FontSettings to LoadOptions so that fonts are resolved while loading.
        LoadOptions loadOptions = new LoadOptions { FontSettings = usbFontSettings };

        // Load the source document using the configured LoadOptions.
        // If the input file does not exist, create a new empty document instead.
        Document doc;
        const string inputPath = "input.docx";
        if (File.Exists(inputPath))
        {
            doc = new Document(inputPath, loadOptions);
        }
        else
        {
            doc = new Document();
        }

        // Add some text that uses special symbols to verify font loading.
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Use a font that is likely to be present; if a custom font from the USB folder is needed,
        // replace the name with the exact font name.
        builder.Font.Name = "Arial";
        builder.Writeln("Special symbols: Ω, 漢字, 😊");

        // Save the processed document. The fonts will be resolved from the USB drive if needed.
        // Save as DOCX to avoid PDF dependencies in environments without the PDF add‑on.
        doc.Save("output.docx");
    }
}
