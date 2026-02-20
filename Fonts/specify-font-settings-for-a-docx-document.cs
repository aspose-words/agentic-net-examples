using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;

class FontSettingsExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize FontSettings and assign it to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Specify a folder that contains TrueType fonts.
        // The folder path can be adjusted to point to any directory with .ttf files.
        string fontsFolder = Path.Combine(Environment.CurrentDirectory, "MyFonts");
        FolderFontSource folderSource = new FolderFontSource(fontsFolder, false);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

        // Build automatic fallback settings based on the fonts available in the folder.
        FontFallbackSettings fallback = fontSettings.FallbackSettings;
        fallback.BuildAutomatic();

        // Configure a simple substitution rule: if "Times New Roman" is missing, use "Arial".
        // This uses the table substitution mechanism.
        fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Times New Roman", new[] { "Arial" });

        // Use DocumentBuilder to write some text with a font that does not exist in the folder.
        // The fallback/substitution scheme will be applied when rendering this text.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This paragraph uses a missing font and will trigger fallback/substitution.");

        // Save the document as DOCX using OoxmlSaveOptions.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
        saveOptions.UpdateLastSavedTimeProperty = true;
        doc.Save("FontSettingsExample.docx", saveOptions);
    }
}
