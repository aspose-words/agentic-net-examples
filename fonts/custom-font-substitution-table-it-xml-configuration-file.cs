using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create temporary working directories.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeExample");
        Directory.CreateDirectory(workDir);
        string fontsDir = Path.Combine(workDir, "MyFonts");
        Directory.CreateDirectory(fontsDir);
        string outPath = Path.Combine(workDir, "Result.pdf");

        // Create a new empty document.
        Document doc = new Document();

        // Create FontSettings and assign them to the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Restrict Aspose.Words to look for fonts only in the custom folder (which is empty in this example).
        FolderFontSource folderSource = new FolderFontSource(fontsDir, false);
        fontSettings.SetFontsSources(new FontSourceBase[] { folderSource });

        // Load the custom font substitution table from an XML string.
        string xml = @"
<SubstitutionTable>
    <Substitution>
        <SourceFont>Arial</SourceFont>
        <TargetFont>Times New Roman</TargetFont>
    </Substitution>
</SubstitutionTable>";

        // Write the XML to a temporary file because TableSubstitutionRule only supports loading from a file.
        string tempXmlPath = Path.Combine(workDir, "substitution.xml");
        File.WriteAllText(tempXmlPath, xml);

        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.Load(tempXmlPath);

        // Write some text using a font that is not present in the custom folder.
        // The substitution rule loaded from XML will determine which font is actually used.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // Assume Arial is missing from MyFonts.
        builder.Writeln("This text should be rendered with the substitute defined in the XML.");

        // Save the resulting document.
        doc.Save(outPath);
        Console.WriteLine($"Document saved to: {outPath}");
    }
}
