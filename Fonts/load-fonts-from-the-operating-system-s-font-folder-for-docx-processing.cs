using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Retrieve all system font folders available on the current OS.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();

        // Create a FontSettings instance and configure it to search the retrieved folders.
        // The second argument (true) enables recursive scanning of subfolders.
        FontSettings fontSettings = new FontSettings();
        fontSettings.SetFontsFolders(systemFontFolders, true);

        // Apply the FontSettings to LoadOptions so they are used while loading the document.
        LoadOptions loadOptions = new LoadOptions
        {
            FontSettings = fontSettings
        };

        // Load a DOCX document using the specified font settings.
        Document doc = new Document("input.docx", loadOptions);

        // Save the processed document (e.g., as PDF) – this will use the loaded fonts.
        doc.Save("output.pdf");
    }
}
