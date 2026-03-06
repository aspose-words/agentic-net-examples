using System;
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

        // First paragraph uses a font that is normally available.
        builder.Font.Name = "Arial";
        builder.Writeln("Hello world!");

        // Second paragraph uses a font that may be unavailable on the system.
        builder.Font.Name = "Amethysta";
        builder.Writeln("The quick brown fox jumps over the lazy dog.");

        // Configure font substitution for the unavailable font.
        // If "Amethysta" cannot be found, Aspose.Words will try "Arvo" first,
        // and if that is also missing, it will fall back to "Courier New".
        doc.FontSettings = new FontSettings();
        doc.FontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "Amethysta", new[] { "Arvo", "Courier New" });

        // Save the document to a DOCX file.
        doc.Save("FontSettingsExample.docx");
    }
}
