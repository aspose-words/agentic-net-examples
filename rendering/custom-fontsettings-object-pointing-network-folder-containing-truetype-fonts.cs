using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the network folder that contains TrueType fonts.
        // Example UNC path: \\ServerName\SharedFolder\Fonts
        string networkFontFolder = @"\\ServerName\SharedFolder\Fonts";

        // Set to true if you want Aspose.Words to scan subfolders recursively.
        bool scanSubfolders = true;

        // Create a custom FontSettings instance.
        FontSettings customFontSettings = new FontSettings();

        // Configure the FontSettings to use the network folder as the font source.
        // This uses the SetFontsFolder method as defined in the Aspose.Words API.
        customFontSettings.SetFontsFolder(networkFontFolder, scanSubfolders);

        // The FontSettings object is now ready to be assigned to a Document
        // or to LoadOptions when loading a document, e.g.:
        // Document doc = new Document("input.docx", new LoadOptions { FontSettings = customFontSettings });
        // (Document creation/loading/saving is omitted as per the task requirements.)
    }
}
