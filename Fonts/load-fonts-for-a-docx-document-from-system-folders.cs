using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Load the DOCX document.
        Document doc = new Document("Input.docx");

        // Retrieve all system font folders.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();

        // Create a FolderFontSource for each system folder (non‑recursive scan).
        FontSourceBase[] folderSources = systemFontFolders
            .Select(folder => new FolderFontSource(folder, false) as FontSourceBase)
            .ToArray();

        // Create a new FontSettings instance and assign the folder sources.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsSources(folderSources);

        // Apply the FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document with the loaded fonts.
        doc.Save("Output.docx");
    }
}
