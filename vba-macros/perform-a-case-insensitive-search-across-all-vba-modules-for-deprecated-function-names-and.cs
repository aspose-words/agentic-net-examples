using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file paths.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string originalPath = Path.Combine(artifactsDir, "Original.docm");
        string updatedPath = Path.Combine(artifactsDir, "Updated.docm");

        // -----------------------------------------------------------------
        // 1. Create a macro‑enabled document with a VBA project and sample modules.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // Create a new VBA project if none exists.
        VbaProject project = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = project;

        // Sample VBA module 1.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub Test()
    Call OldFunc
    Call DeprecatedMethod
End Sub

Function OldFunc() As Integer
    OldFunc = 1
End Function
"
        };
        doc.VbaProject.Modules.Add(module1);

        // Sample VBA module 2 (empty source to test null handling).
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = null // Intentionally null.
        };
        doc.VbaProject.Modules.Add(module2);

        // Save the document in macro‑enabled format.
        doc.Save(originalPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 2. Load the document and perform case‑insensitive replacement of deprecated functions.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalPath);

        // Ensure the document actually contains a VBA project.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            // Define deprecated function names and their replacements.
            var replacements = new (string oldName, string newName)[]
            {
                ("OldFunc", "NewFunc"),
                ("DeprecatedMethod", "ModernMethod")
            };

            foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = vbaModule.SourceCode ?? string.Empty;
                string updatedSource = source;

                foreach (var (oldName, newName) in replacements)
                {
                    // Use word boundaries to avoid partial matches.
                    string pattern = $@"\b{Regex.Escape(oldName)}\b";
                    updatedSource = Regex.Replace(updatedSource, pattern, newName, RegexOptions.IgnoreCase);
                }

                // Apply the updated source back to the module.
                vbaModule.SourceCode = updatedSource;
            }

            // Save the modified document.
            loadedDoc.Save(updatedPath, SaveFormat.Docm);
        }

        // Indicate completion (no interactive input).
        Console.WriteLine("VBA modules processed and saved to:");
        Console.WriteLine(updatedPath);
    }
}
