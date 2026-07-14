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

        // Paths for the XML substitution file and the rendered PDF.
        string substitutionXmlPath = Path.Combine(artifactsDir, "fontSubstitution.xml");
        string outputPdfPath = Path.Combine(artifactsDir, "output.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample document that uses a font which is not available.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text is formatted with a missing font and will be substituted.");

        // -----------------------------------------------------------------
        // 2. Define a font substitution rule and save it to an XML file.
        // -----------------------------------------------------------------
        FontSettings fontSettingsForSaving = new FontSettings();
        var tableRuleForSaving = fontSettingsForSaving.SubstitutionSettings.TableSubstitution;
        // Substitute the missing font with a common system font (e.g., Arial).
        tableRuleForSaving.AddSubstitutes("MissingFont", "Arial");
        // Save the substitution table to XML.
        tableRuleForSaving.Save(substitutionXmlPath);

        // -----------------------------------------------------------------
        // 3. Load the substitution rules from the XML file.
        // -----------------------------------------------------------------
        FontSettings loadedFontSettings = new FontSettings();
        var tableRuleForLoading = loadedFontSettings.SubstitutionSettings.TableSubstitution;
        tableRuleForLoading.Load(substitutionXmlPath);

        // Apply the loaded font settings to the document.
        doc.FontSettings = loadedFontSettings;

        // -----------------------------------------------------------------
        // 4. Render the document to PDF using the loaded substitution rules.
        // -----------------------------------------------------------------
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // -----------------------------------------------------------------
        // 5. Verify that the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // The program finishes without requiring any user interaction.
    }
}
