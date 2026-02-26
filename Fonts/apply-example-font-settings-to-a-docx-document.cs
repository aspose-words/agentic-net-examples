using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write some text using two different fonts.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");
        builder.Font.Name = "Amethysta";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Set up font substitution: if "Amethysta" is unavailable,
        // first try "Arvo", then "Courier New".
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Amethysta", new[] { "Arvo", "Courier New" });
        doc.FontSettings = fontSettings;

        // Ensure the output directory exists.
        string artifactsDir = Path.Combine(Environment.CurrentDirectory, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Save the document to the output folder.
        doc.Save(Path.Combine(artifactsDir, "FontSettingsExample.docx"));
    }
}
