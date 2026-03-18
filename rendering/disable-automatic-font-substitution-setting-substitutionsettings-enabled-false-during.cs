using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a FontSettings instance to configure font handling.
        FontSettings fontSettings = new FontSettings();

        // Disable all built‑in font substitution rules.
        fontSettings.SubstitutionSettings.FontNameSubstitution.Enabled = false;      // Font name rule
        fontSettings.SubstitutionSettings.FontInfoSubstitution.Enabled = false;      // Font info rule
        fontSettings.SubstitutionSettings.TableSubstitution.Enabled = false;         // Table substitution rule
        fontSettings.SubstitutionSettings.FontConfigSubstitution.Enabled = false;    // Font config rule
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.Enabled = false;   // Default font rule

        // Apply the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // Add some text that uses a font which is not available.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This text uses a missing font and will not be substituted.");

        // Render the document to PDF (or any other supported format).
        doc.Save("Output.pdf");
    }
}
