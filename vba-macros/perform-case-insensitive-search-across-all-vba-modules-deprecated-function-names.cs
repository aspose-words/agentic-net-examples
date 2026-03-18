using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleUpdater
{
    static void Main()
    {
        // Load the document that contains VBA macros.
        string inputPath = "input.docm";
        Document doc = new Document(inputPath);

        // Map of deprecated function names (key) to their replacements (value).
        var deprecatedFunctions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "OldFunc1", "NewFunc1" },
            { "LegacyCalc", "ModernCalc" },
            // Add more pairs as needed.
        };

        // Iterate through all VBA modules in the project.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            string source = module.SourceCode;

            // Replace each deprecated function name with its new name, case‑insensitively.
            foreach (var kvp in deprecatedFunctions)
            {
                // Escape the key to treat it as a literal pattern.
                string pattern = Regex.Escape(kvp.Key);
                source = Regex.Replace(source, pattern, kvp.Value, RegexOptions.IgnoreCase);
            }

            // Write the updated source back to the module.
            module.SourceCode = source;
        }

        // Save the modified document.
        string outputPath = "output.docm";
        doc.Save(outputPath);
    }
}
