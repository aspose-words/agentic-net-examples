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

        // Write a line using a font that is available on most systems.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");

        // Write a line using a font that is likely unavailable.
        builder.Font.Name = "Amethysta";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Set up font substitution rules.
        // If "Amethysta" cannot be found, Aspose.Words will try "Arvo" first,
        // and if that is also missing, it will fall back to "Courier New".
        FontSettings fontSettings = new FontSettings();
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Amethysta", new[] { "Arvo", "Courier New" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document to a DOCX file.
        string outputPath = "FontSettingsExample.docx";
        doc.Save(outputPath);
    }
}
