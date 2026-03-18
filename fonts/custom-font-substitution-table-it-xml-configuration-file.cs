using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Define paths to the data folder, custom fonts folder, XML substitution file and output PDF.
        string dataDir = @"C:\Data\";
        string fontsDir = Path.Combine(dataDir, "MyFonts");
        string xmlPath = Path.Combine(dataDir, "FontSubstitutionRules.xml");
        string outPath = Path.Combine(dataDir, "Result.pdf");

        // Create a new empty document.
        Document doc = new Document();

        // Create FontSettings and assign them to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Restrict Aspose.Words to look for fonts only in the custom folder.
        FolderFontSource folderSource = new FolderFontSource(fontsDir, false);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

        // Load the custom font substitution table from the XML configuration file.
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.Load(xmlPath);

        // Write some text using a font that is not present in the custom folder.
        // The substitution rule loaded from XML will determine which font is actually used.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Assume Arial is missing from MyFonts.
        builder.Writeln("This text should be rendered with the substitute defined in the XML.");

        // Save the resulting document.
        doc.Save(outPath);
    }
}
