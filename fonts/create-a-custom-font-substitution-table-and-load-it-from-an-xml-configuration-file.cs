using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directories and file names.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(artifactsDir);
        string substitutionFile = Path.Combine(artifactsDir, "custom_substitution.xml");
        string outputDoc = Path.Combine(artifactsDir, "CustomFontSubstitution.docx");

        // Create a new document and add a paragraph using a font that likely does not exist.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Configure FontSettings for the document.
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Load the default Windows substitution table and save it to an XML file.
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.LoadWindowsSettings();
        tableRule.Save(substitutionFile);

        // Reload the substitution table from the XML file (demonstrates loading from configuration).
        tableRule.Load(substitutionFile);

        // Add a custom substitution: replace "MissingFont" with "Arial" or "Courier New" if needed.
        tableRule.AddSubstitutes("MissingFont", "Arial", "Courier New");

        // Save the document; Aspose.Words will apply the substitution rules during rendering.
        doc.Save(outputDoc);
    }
}
