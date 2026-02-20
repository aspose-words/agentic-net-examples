using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

class LoadSystemFontsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Assign default FontSettings to the document.
        doc.FontSettings = new FontSettings();

        // By default a blank document contains a SystemFontSource.
        // Retrieve it from the document's font sources.
        FontSourceBase[] sources = doc.FontSettings.GetFontsSources();
        SystemFontSource systemFontSource = sources.OfType<SystemFontSource>().FirstOrDefault();

        if (systemFontSource == null)
        {
            Console.WriteLine("System font source not found.");
            return;
        }

        // List all system font folders detected by Aspose.Words.
        Console.WriteLine("System font folders:");
        foreach (string folder in SystemFontSource.GetSystemFontFolders())
        {
            Console.WriteLine($"  {folder}");
        }

        // List all fonts available via the system font source.
        Console.WriteLine("\nAvailable system fonts:");
        foreach (PhysicalFontInfo fontInfo in systemFontSource.GetAvailableFonts())
        {
            Console.WriteLine($"  Family: {fontInfo.FontFamilyName}, Full Name: {fontInfo.FullFontName}");
        }

        // Example: set a substitution for a missing font.
        doc.FontSettings.SubstitutionSettings.TableSubstitution.Enabled = true;
        doc.FontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("MissingFont", "Arial", "Times New Roman");

        // No saving is required for this example, but the document can be saved normally.
        // doc.Save("Output.docx");
    }
}
