using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Retrieve all system font folders.
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();

        // Convert each folder into a FolderFontSource (scan subfolders recursively).
        FontSourceBase[] fontSources = systemFontFolders
            .Select(folder => new FolderFontSource(folder, true) as FontSourceBase)
            .ToArray();

        // Assign the font sources to the document's FontSettings.
        doc.FontSettings = new FontSettings();
        doc.FontSettings.SetFontsSources(fontSources);

        // Add some sample text using a common system font.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This text is rendered with the Arial font from the system fonts.");

        // Save the document.
        doc.Save("Output.docx");
    }
}
