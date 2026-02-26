using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontSettingsExample
{
    static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document.
        Document doc = new Document();

        // Access the document builder to insert text.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Set default font for the whole document (applies to new styles).
        // -----------------------------------------------------------------
        doc.Styles.DefaultFont.Name = "Arial";
        doc.Styles.DefaultFont.Size = 12;

        // -----------------------------------------------------------------
        // 2. Configure font substitution rules.
        //    If a font is missing, substitute it with the listed alternatives.
        // -----------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();

        // Example: substitute the unavailable font "Amethysta" with "Arvo" then "Courier New".
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Amethysta", new[] { "Arvo", "Courier New" });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // -----------------------------------------------------------------
        // 3. Write text using both available and unavailable fonts.
        // -----------------------------------------------------------------
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses the default Arial font.");

        builder.Font.Name = "Amethysta"; // This font is likely missing.
        builder.Writeln("This line uses Amethysta, which will be substituted.");

        // -----------------------------------------------------------------
        // 4. Save the document as DOCX.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(artifactsDir, "FontSettingsExample.docx");
        doc.Save(outputPath);
    }
}
