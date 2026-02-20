using System;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Replacing;

namespace FontSubstitutionExample
{
    class Program
    {
        static void Main()
        {
            // Load a DOCX document that uses a font which may be missing on the system.
            Document doc = new Document("Input.docx"); // load

            // Collect warnings that occur during loading/saving (e.g., font substitution warnings).
            WarningInfoCollection warnings = new WarningInfoCollection();
            doc.WarningCallback = warnings;

            // Configure font settings.
            FontSettings fontSettings = new FontSettings();

            // Set a default font that will be used when no other substitution rule resolves the missing font.
            fontSettings.SubstitutionSettings.DefaultFontSubstitution.DefaultFontName = "Arial";

            // Optionally, point Aspose.Words to a folder that contains additional TrueType fonts.
            // Replace "FontsFolderPath" with the actual path to your fonts directory.
            fontSettings.SetFontsFolder("FontsFolderPath", false);

            // Add a table substitution: if "Arial" is not found, use "Arvo" or "Slab" as substitutes.
            fontSettings.SubstitutionSettings.TableSubstitution.AddSubstitutes("Arial", "Arvo", "Slab");

            // Apply the configured font settings to the document.
            doc.FontSettings = fontSettings;

            // Save the document to PDF (or any other format). This triggers the font substitution process.
            doc.Save("Output.pdf"); // save

            // Output any font substitution warnings that were captured.
            foreach (WarningInfo info in warnings)
            {
                if (info.WarningType == WarningType.FontSubstitution)
                {
                    Console.WriteLine(info.Description);
                }
            }
        }
    }
}
