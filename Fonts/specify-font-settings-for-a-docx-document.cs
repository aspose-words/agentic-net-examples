using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class FontSettingsExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add text with different fonts.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial";
        builder.Writeln("This line uses the Arial font.");

        // This font does not exist on the system; it will be substituted.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Create a FontSettings object to control font resolution.
        FontSettings fontSettings = new FontSettings();

        // Optional: add a custom folder that contains additional fonts.
        // Replace "MyFonts" with the actual path to your font directory.
        string customFontsFolder = Path.Combine(Environment.CurrentDirectory, "MyFonts");
        if (Directory.Exists(customFontsFolder))
            fontSettings.SetFontsFolder(customFontsFolder, false);

        // Define a substitution rule: when "MissingFont" cannot be found,
        // use "Courier New" as the first substitute.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "MissingFont", new[] { "Courier New" });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to a DOCX file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "FontSettingsExample.docx");
        doc.Save(outputPath);
    }
}
