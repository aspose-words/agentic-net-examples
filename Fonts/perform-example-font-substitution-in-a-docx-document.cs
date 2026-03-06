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

        // Write text using a font that does not exist on the system.
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Collect warnings generated during processing.
        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.WarningCallback = warnings;

        // Configure font substitution settings.
        FontSettings fontSettings = new FontSettings();

        // Table substitution: replace "MissingFont" with "Arial", then "Times New Roman".
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "MissingFont", new[] { "Arial", "Times New Roman" });

        // Default substitution rule: use "Courier New" if no other substitute is found.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Courier New";

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document.
        doc.Save("FontSubstitutionExample.docx");

        // Output any font substitution warnings.
        foreach (WarningInfo info in warnings)
        {
            if (info.WarningType == WarningType.FontSubstitution)
            {
                FontSubstitutionWarningInfo fsInfo = (FontSubstitutionWarningInfo)info;
                Console.WriteLine($"Requested: {fsInfo.RequestedFamilyName}, " +
                                  $"Resolved: {fsInfo.ResolvedFont?.FullFontName}, " +
                                  $"Reason: {fsInfo.Reason}");
            }
        }
    }
}
