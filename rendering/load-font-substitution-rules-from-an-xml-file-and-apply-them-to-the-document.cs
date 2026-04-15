using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Define folders for output artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample document that uses a font which is likely missing.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont"; // Font that probably does not exist on the system.
        builder.Writeln("This text uses a missing font and will be substituted.");

        // -----------------------------------------------------------------
        // 2. Prepare a font substitution table and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;

        // Define a substitute: map "MissingFont" to a common system font "Arial".
        tableRule.AddSubstitutes("MissingFont", "Arial");

        // Save the substitution table to XML.
        string substitutionXmlPath = Path.Combine(artifactsDir, "FontSubstitutionRules.xml");
        tableRule.Save(substitutionXmlPath);

        // -----------------------------------------------------------------
        // 3. Load the substitution rules from the XML file and apply them.
        // -----------------------------------------------------------------
        // Create a fresh FontSettings instance to demonstrate loading.
        FontSettings loadedFontSettings = new FontSettings();
        TableSubstitutionRule loadedTableRule = loadedFontSettings.SubstitutionSettings.TableSubstitution;
        loadedTableRule.Load(substitutionXmlPath);

        // Assign the loaded settings to the document.
        doc.FontSettings = loadedFontSettings;

        // -----------------------------------------------------------------
        // 4. Render the document to PDF using the substitution rules.
        // -----------------------------------------------------------------
        string pdfPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(pdfPath);

        // -----------------------------------------------------------------
        // 5. Validate that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output was not created.");

        // The example finishes without requiring user interaction.
    }
}
