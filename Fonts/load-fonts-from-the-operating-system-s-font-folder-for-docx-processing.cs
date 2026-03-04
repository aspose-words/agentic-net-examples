using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

class LoadSystemFontsExample
{
    static void Main()
    {
        // Create a FontSettings instance that will hold the font sources.
        FontSettings fontSettings = new FontSettings();

        // Retrieve all system font folders available on the current OS.
        string[] systemFolders = SystemFontSource.GetSystemFontFolders();

        // Convert each folder path into a FolderFontSource (non‑recursive scan).
        FontSourceBase[] folderSources = systemFolders
            .Select(path => new FolderFontSource(path, false))
            .Cast<FontSourceBase>()
            .ToArray();

        // Combine the default system font source with the folder sources.
        // The SystemFontSource represents all fonts installed on the system.
        FontSourceBase[] allSources = new FontSourceBase[folderSources.Length + 1];
        allSources[0] = new SystemFontSource(); // system source
        Array.Copy(folderSources, 0, allSources, 1, folderSources.Length);

        // Apply the combined font sources to the FontSettings object.
        fontSettings.SetFontsSources(allSources);

        // Load the document without LoadOptions (LoadOptions is not required for font settings).
        Document doc = new Document("Input.docx");

        // Assign the prepared FontSettings to the loaded document.
        doc.FontSettings = fontSettings;

        // Perform any required processing on the document here.
        // ...

        // Save the document to verify that fonts are resolved correctly.
        doc.Save("Output.pdf");
    }
}
