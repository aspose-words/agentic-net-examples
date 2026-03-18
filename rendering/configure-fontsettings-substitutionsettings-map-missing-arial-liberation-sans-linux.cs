using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Fonts;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a FontSettings instance to hold substitution rules.
        FontSettings fontSettings = new FontSettings();

        // Determine whether the current OS is Linux or macOS.
        bool isLinuxOrMac = new[] { PlatformID.Unix, PlatformID.MacOSX }
            .Any(p => Environment.OSVersion.Platform == p);

        if (isLinuxOrMac)
        {
            // On Linux/macOS use the table substitution rule to map missing "Arial"
            // to the available "Liberation Sans" font.
            TableSubstitutionRule tableRule = fontSettings.SubstitutionSettings.TableSubstitution;
            tableRule.SetSubstitutes("Arial", "Liberation Sans");
        }

        // Apply the configured FontSettings to the document.
        doc.FontSettings = fontSettings;

        // The document can now be processed (rendered, saved, etc.) with the substitution in effect.
    }
}
