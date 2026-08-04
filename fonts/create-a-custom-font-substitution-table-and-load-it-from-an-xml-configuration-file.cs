using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // Step 1: Create a font substitution table using the built‑in Windows settings
        // and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings initialFontSettings = new FontSettings();
        TableSubstitutionRule initialTable = initialFontSettings.SubstitutionSettings.TableSubstitution;

        // Load the default Windows substitution table.
        initialTable.LoadWindowsSettings();

        // Save the table to a custom XML file.
        string substitutionXmlPath = Path.Combine(artifactsDir, "CustomFontSubstitution.xml");
        initialTable.Save(substitutionXmlPath);

        // -----------------------------------------------------------------
        // Step 2: Load the custom substitution table from the XML file
        // and add an extra substitute for a font that does not exist.
        // -----------------------------------------------------------------
        FontSettings customFontSettings = new FontSettings();
        TableSubstitutionRule customTable = customFontSettings.SubstitutionSettings.TableSubstitution;

        // Load the previously saved XML file.
        customTable.Load(substitutionXmlPath);

        // Add a substitute chain for a missing font.
        // If "MissingFont" is not found, Aspose.Words will try "Arial" first,
        // then "Courier New" if "Arial" is also unavailable.
        customTable.AddSubstitutes("MissingFont", "Arial", "Courier New");

        // -----------------------------------------------------------------
        // Step 3: Create a document that uses the missing font and apply the
        // custom FontSettings.
        // -----------------------------------------------------------------
        Document doc = new Document();
        doc.FontSettings = customFontSettings;

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line is formatted with a font that does not exist.");
        builder.Writeln("It will be rendered using the substitution chain defined in the XML file.");

        // Save the resulting document.
        string outputPath = Path.Combine(artifactsDir, "DocumentWithCustomSubstitution.pdf");
        doc.Save(outputPath);
    }
}
