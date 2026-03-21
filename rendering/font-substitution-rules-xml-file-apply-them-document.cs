using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Use the system fonts folder as the custom fonts source.
        string fontsDir = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

        // Create a temporary directory for the XML file and output.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeExample");
        Directory.CreateDirectory(tempDir);

        // Path to the XML file that defines font substitution rules.
        string substitutionXml = Path.Combine(tempDir, "FontSubstitutionRules.xml");

        // Write a minimal substitution rule XML file.
        File.WriteAllText(substitutionXml,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<SubstitutionTable>
  <Substitution>
    <Source>Arial</Source>
    <Target>Times New Roman</Target>
  </Substitution>
</SubstitutionTable>");

        // Folder where the output document will be saved.
        string artifactsDir = Path.Combine(tempDir, "Output");
        Directory.CreateDirectory(artifactsDir);

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
        string outputPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
