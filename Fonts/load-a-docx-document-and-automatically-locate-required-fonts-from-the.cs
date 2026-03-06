using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = @"C:\Docs\Input.docx";
        string outputPath = @"C:\Docs\Output.docx";

        // Create LoadOptions and configure FontSettings to use system font folders
        LoadOptions loadOptions = new LoadOptions();
        FontSettings fontSettings = new FontSettings();

        // Get all system font directories
        string[] systemFontFolders = SystemFontSource.GetSystemFontFolders();

        // Add those directories as font sources (search recursively)
        fontSettings.SetFontsFolders(systemFontFolders, true);

        // Assign the FontSettings to the LoadOptions
        loadOptions.FontSettings = fontSettings;

        // Load the document using the configured LoadOptions
        Document doc = new Document(inputPath, loadOptions);

        // Save the document (can be the same format or another, e.g., PDF)
        doc.Save(outputPath);
    }
}
