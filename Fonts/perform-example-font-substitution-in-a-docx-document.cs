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

        // Add some text that uses a font which is not installed on the system.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "MissingFont";
        builder.Writeln("This line uses a missing font and will be substituted.");

        // Collect warnings that occur during loading/saving.
        WarningInfoCollection warnings = new WarningInfoCollection();
        doc.WarningCallback = warnings;

        // Configure font substitution settings.
        FontSettings fontSettings = new FontSettings();

        // Set a default font to be used when no other substitute is found.
        fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

        // Define a substitution table: try "Times New Roman", then "Courier New" for "MissingFont".
        fontSettings.SubstitutionSettings.TableSubstitution.SetSubstitutes(
            "MissingFont", "Times New Roman", "Courier New");

        // Apply the font settings to the document.
        doc.FontSettings = fontSettings;

        // Save the document (creates a DOCX file).
        doc.Save("FontSubstitutionExample.docx");

        // Output any font substitution warnings.
        foreach (WarningInfo info in warnings)
        {
            if (info.WarningType == WarningType.FontSubstitution)
            {
                FontSubstitutionWarningInfo fontInfo = (FontSubstitutionWarningInfo)info;
                Console.WriteLine($"Requested: '{fontInfo.RequestedFamilyName}' -> Resolved: '{fontInfo.ResolvedFont.FullFontName}' (Reason: {fontInfo.Reason})");
            }
        }
    }
}
