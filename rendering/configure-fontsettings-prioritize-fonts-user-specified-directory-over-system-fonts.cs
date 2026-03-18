using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontPriorityExample
{
    static void Main()
    {
        // Path to the user‑specified folder that contains TrueType fonts.
        string userFontsFolder = @"C:\MyFonts";

        // Create a folder font source for the user folder.
        // The third argument sets the priority – a higher value means higher priority.
        FolderFontSource userFontSource = new FolderFontSource(userFontsFolder, true, 1);

        // Obtain the default system font source(s). By default there is one SystemFontSource.
        FontSourceBase[] systemSources = FontSettings.DefaultInstance.GetFontsSources();

        // Combine the user font source with the system sources, placing the user source first.
        FontSourceBase[] combinedSources = new FontSourceBase[systemSources.Length + 1];
        combinedSources[0] = userFontSource;
        Array.Copy(systemSources, 0, combinedSources, 1, systemSources.Length);

        // Create a FontSettings instance and assign the combined sources.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsSources(combinedSources);

        // Apply the FontSettings while loading the document.
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.FontSettings = fontSettings;

        // Load the source document (replace with your actual file path).
        Document doc = new Document(@"C:\Docs\Input.docx", loadOptions);

        // Render the document – the user‑specified fonts will be used before system fonts.
        doc.Save(@"C:\Docs\Output.pdf");
    }
}
