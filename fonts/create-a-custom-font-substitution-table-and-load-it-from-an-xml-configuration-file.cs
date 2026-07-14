using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the custom substitution table XML.
        string substitutionXmlPath = Path.Combine(artifactsDir, "CustomFontSubstitution.xml");

        // -----------------------------------------------------------------
        // Step 1: Create a font substitution table and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings tempSettings = new FontSettings();
        TableSubstitutionRule tempTable = tempSettings.SubstitutionSettings.TableSubstitution;

        // Define substitutes for a font that does not exist in the system.
        // When "MissingFont" is requested, Aspose.Words will try "Arial" first,
        // then "Courier New" if the first substitute is unavailable.
        tempTable.SetSubstitutes("MissingFont", new[] { "Arial", "Courier New" });

        // Save the substitution table to an XML file.
        tempTable.Save(substitutionXmlPath);

        // -----------------------------------------------------------------
        // Step 2: Load the custom substitution table from the XML file.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
        tableRule.Load(substitutionXmlPath);

        // -----------------------------------------------------------------
        // Step 3: Create a document that uses the missing font.
        // -----------------------------------------------------------------
        Document doc = new Document();
        doc.FontSettings = fontSettings;

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted according to the custom table.");

        // -----------------------------------------------------------------
        // Step 4: Save the document.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "CustomFontSubstitution.pdf");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully to: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
