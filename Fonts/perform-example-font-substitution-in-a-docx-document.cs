using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Loading;

class FontSubstitutionExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write text using a font that does not exist on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Collect warnings that occur during loading/saving.
        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.WarningCallback = warnings;

        // Configure font substitution settings.
        FontSettings fontSettings = new FontSettings();

        // Default substitution rule – use Arial when no other substitute is found.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Table substitution rule – try these fonts in order when "MissingFont" is unavailable.
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "MissingFont", new[] { "Times New Roman", "Courier New" });

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document (triggers font substitution).
        doc.Save("FontSubstitutionResult.docx");

        // Output any font substitution warnings.
        foreach (WarningInfo info in warnings)
        {
            if (info.WarningType == WarningType.FontSubstitution)
            {
                FontSubstitutionWarningInfo fontInfo = (FontSubstitutionWarningInfo)info;
                Console.WriteLine($"Requested font: {fontInfo.RequestedFamilyName}");
                Console.WriteLine($"Substituted with: {fontInfo.ResolvedFont?.FullFontName}");
                Console.WriteLine($"Reason: {fontInfo.Reason}");
                Console.WriteLine($"Description: {fontInfo.Description}");
            }
        }
    }
}
