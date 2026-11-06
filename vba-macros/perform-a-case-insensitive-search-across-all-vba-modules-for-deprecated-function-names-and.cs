using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add a sample VBA module if none exist.
        VbaModuleCollection modules = doc.VbaProject.Modules;
        if (modules.Count == 0)
        {
            string sampleCode = @"
Sub SampleMacro()
    Call OldFunction()
    Call LegacyCalc(5)
End Sub
";

            // VbaModule has a parameterless constructor in this version.
            VbaModule sampleModule = new VbaModule();
            sampleModule.Name = "Module1";
            sampleModule.SourceCode = sampleCode;
            modules.Add(sampleModule);
        }

        // Define deprecated function names and their replacements (case‑insensitive).
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "OldFunction", "NewFunction" },
            { "LegacyCalc", "ModernCalc" }
        };

        // Perform case‑insensitive search and replace in each module.
        foreach (VbaModule module in modules)
        {
            string source = module.SourceCode ?? string.Empty;

            foreach (var kvp in replacements)
            {
                string pattern = @"\b" + Regex.Escape(kvp.Key) + @"\b";
                source = Regex.Replace(source, pattern, kvp.Value, RegexOptions.IgnoreCase);
            }

            module.SourceCode = source;
        }

        // Save the document in macro‑enabled format.
        const string outputPath = "output.docm";
        doc.Save(outputPath, SaveFormat.Docm);

        // Reload the document to verify changes.
        Document loadedDoc = new Document(outputPath);
        bool allReplaced = true;

        if (loadedDoc.VbaProject != null)
        {
            foreach (VbaModule module in loadedDoc.VbaProject.Modules)
            {
                string source = module.SourceCode ?? string.Empty;
                foreach (var kvp in replacements)
                {
                    if (source.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        allReplaced = false;
                    }
                }
            }
        }

        Console.WriteLine(allReplaced
            ? "All deprecated functions were successfully replaced."
            : "Some deprecated functions were not replaced.");
    }
}
