using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare folders for output and temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string xmlPath = Path.Combine(outputDir, "custom_substitution.xml");
        string docPath = Path.Combine(outputDir, "Result.pdf");

        // Create a new document and a FontSettings instance.
        Document doc = new Document();
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Obtain the table substitution rule.
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;

        // Load the built‑in Windows substitution table and save it to an XML file.
        tableRule.LoadWindowsSettings();
        tableRule.Save(xmlPath);

        // Now load the substitution table back from the XML file.
        tableRule.Load(xmlPath);

        // Add a custom substitution: when the font "MissingFont" is not found,
        // use "Courier New" as a fallback.
        tableRule.AddSubstitutes("MissingFont", "Courier New");

        // Build some content using a font that does not exist on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and should be rendered with the substitute.");

        // Save the document. The output file will be created in the Output folder.
        doc.Save(docPath);
    }
}
