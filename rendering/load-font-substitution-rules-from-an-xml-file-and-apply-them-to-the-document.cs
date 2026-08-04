using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Prepare folders for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path to the temporary XML file that will hold the substitution table.
        string substitutionXmlPath = Path.Combine(artifactsDir, "FontSubstitutionTable.xml");

        // -----------------------------------------------------------------
        // Step 1: Create a substitution rule and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings tempFontSettings = new FontSettings();
        TableSubstitutionRule tempRule = tempFontSettings.SubstitutionSettings.TableSubstitution;

        // Define a substitute: when the document uses "MissingFont", replace it with "Arial".
        tempRule.AddSubstitutes("MissingFont", "Arial");

        // Save the rule to XML.
        tempRule.Save(substitutionXmlPath);

        // -----------------------------------------------------------------
        // Step 2: Create a document that uses a font which is not available.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line is written with a missing font and should be substituted.");

        // -----------------------------------------------------------------
        // Step 3: Load the substitution rules from the XML file and apply them.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();
        TableSubstitutionRule rule = fontSettings.SubstitutionSettings.TableSubstitution;
        rule.Load(substitutionXmlPath);
        doc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // Step 4: Render the document to PDF (substitution will be applied).
        // -----------------------------------------------------------------
        string outputPdfPath = Path.Combine(artifactsDir, "Result.pdf");
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // Step 5: Verify that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // The program finishes here without requiring any user interaction.
    }
}
