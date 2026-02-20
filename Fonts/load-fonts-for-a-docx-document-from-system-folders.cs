using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class LoadSystemFontsExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Ensure the document has FontSettings (creates a default instance if null).
        if (doc.FontSettings == null)
            doc.FontSettings = new FontSettings();

        // Retrieve the system font source that represents all TrueType fonts installed on the OS.
        FontSourceBase[] sources = doc.FontSettings.GetFontsSources();
        SystemFontSource systemFontSource = null;

        // Find the SystemFontSource among the existing sources.
        foreach (FontSourceBase source in sources)
        {
            if (source is SystemFontSource)
            {
                systemFontSource = (SystemFontSource)source;
                break;
            }
        }

        // If for some reason the system source is missing, create and add it.
        if (systemFontSource == null)
        {
            systemFontSource = new SystemFontSource();
            doc.FontSettings.SetFontsSources(new FontSourceBase[] { systemFontSource });
        }

        // Example: list all available system fonts to the console.
        foreach (PhysicalFontInfo fontInfo in systemFontSource.GetAvailableFonts())
        {
            Console.WriteLine($"FontFamilyName: {fontInfo.FontFamilyName}");
            Console.WriteLine($"FullFontName  : {fontInfo.FullFontName}");
            Console.WriteLine($"Version       : {fontInfo.Version}");
            Console.WriteLine($"FilePath      : {fontInfo.FilePath}");
            Console.WriteLine();
        }

        // Save the document (fonts are now loaded from system folders).
        doc.Save("OutputDocument.docx");
    }
}
