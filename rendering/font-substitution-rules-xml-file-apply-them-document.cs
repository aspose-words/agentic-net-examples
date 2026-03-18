using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Path to the folder that contains custom fonts.
        string fontsDir = @"C:\MyFonts";

        // Path to the XML file that defines font substitution rules.
        string substitutionXml = @"C:\MyFonts\Font substitution rules.xml";

        // Folder where the output document will be saved.
        string artifactsDir = @"C:\Output";

        // Create a new empty document.
        Document doc = new Document();

        // Create FontSettings and assign them to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Use only the custom font folder as a source.
        FolderFontSource folderFontSource = new FolderFontSource(fontsDir, false);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderFontSource });

        // Load the custom table substitution rules from the XML file.
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.Load(substitutionXml); // Loads settings from the specified XML file.

        // Write some text using a font that is not present in the custom folder.
        // The loaded substitution rules will determine which available font is used instead.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Assume Arial is missing from the custom folder.
        builder.Writeln("Text written in Arial, will be substituted according to the loaded rules.");

        // Save the resulting document.
        doc.Save(Path.Combine(artifactsDir, "Result.pdf"));
    }
}
