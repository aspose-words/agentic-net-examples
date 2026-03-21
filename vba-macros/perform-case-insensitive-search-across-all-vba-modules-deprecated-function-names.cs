using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

class VbaModuleUpdater
{
    static void Main()
    {
        // Paths for input and output documents.
        string inputPath = "input.docm";
        string outputPath = "output.docm";

        // Verify that the input file exists.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file '{inputPath}' not found.");
            return;
        }

        // Load the document that contains VBA macros.
        Document doc = new Document(inputPath);

        // If the document has no VBA project, nothing to update.
        if (doc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found in the document.");
            doc.Save(outputPath);
            return;
        }

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
            if (module == null) continue;

            string source = module.SourceCode ?? string.Empty;

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
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
