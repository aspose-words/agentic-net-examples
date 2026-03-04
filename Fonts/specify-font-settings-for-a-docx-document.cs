using System;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line using a font that is guaranteed to exist.
        builder.Font.Name = "Arial";
        builder.Font.Size = 14;
        builder.Writeln("This line uses the Arial font.");

        // Write a line using a font that is likely missing on the system.
        builder.Font.Name = "Amethysta";
        builder.Font.Size = 14;
        builder.Writeln("This line uses the Amethysta font, which will be substituted.");

        // ------------------------------------------------------------
        // Configure document‑wide font settings.
        // ------------------------------------------------------------
        FontSettings fontSettings = new FontSettings();

        // Example: add a custom folder that contains additional fonts.
        // (Uncomment and set the correct path if you have a folder with fonts.)
        // FolderFontSource customFolder = new FolderFontSource(@"C:\MyFonts", true);
        // fontSettings.SetFontsSources(new FontSourceBase[] { customFolder });

        // Define a substitution rule: if "Amethysta" is unavailable,
        // first try "Arvo", then "Courier New".
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Amethysta", new[] { "Arvo", "Courier New" });

        // Assign the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // ------------------------------------------------------------
        // Save the document to a DOCX file.
        // ------------------------------------------------------------
        doc.Save("FontSettingsExample.docx");
    }
}
