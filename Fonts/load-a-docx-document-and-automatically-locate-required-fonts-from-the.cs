using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class Program
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\InputDocument.docx";

        // Create LoadOptions and assign a FontSettings instance.
        // The FontSettings will use the default system font source,
        // which automatically searches the operating system's font folders.
        LoadOptions loadOptions = new LoadOptions();
        FontSettings fontSettings = new FontSettings();

        // Ensure the font sources are reset to the system defaults.
        // This guarantees that Aspose.Words will look for fonts in the
        // system font folders (e.g., C:\Windows\Fonts on Windows).
        fontSettings.ResetFontSources();

        // Attach the FontSettings to the LoadOptions.
        loadOptions.FontSettings = fontSettings;

        // Load the document using the constructor that accepts a file name and LoadOptions.
        Document doc = new Document(inputPath, loadOptions);

        // (Optional) Save the document to verify that it was loaded correctly.
        string outputPath = @"C:\Docs\OutputDocument.docx";
        doc.Save(outputPath);
    }
}
