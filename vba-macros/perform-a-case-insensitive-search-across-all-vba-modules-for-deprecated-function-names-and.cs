using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Path for the initial and updated documents.
        const string initialPath = "Sample.docm";
        const string updatedPath = "Sample_Updated.docm";

        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = vbaProject;

        // 3. Add a procedural module containing deprecated function calls.
        VbaModule module = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub TestMacro()
    Call OldFunction()
    Call LEGACYFUNC()
    MsgBox ""Done""
End Sub

Function OldFunction() As String
    OldFunction = ""Old""
End Function

Function LegacyFunc() As String
    LegacyFunc = ""Legacy""
End Function"
        };
        doc.VbaProject.Modules.Add(module);

        // 4. Save the document in macro‑enabled format.
        doc.Save(initialPath);

        // 5. Load the saved document to simulate a real‑world scenario.
        Document loadedDoc = new Document(initialPath);

        // 6. Define deprecated function names and their replacements.
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "OldFunction", "NewFunction" },
            { "LegacyFunc", "NewFunction" }
        };

        // 7. Iterate over all VBA modules and perform case‑insensitive replacements.
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;

                foreach (var kvp in replacements)
                {
                    // Build a regex pattern that matches the whole word, case‑insensitively.
                    string pattern = $@"\b{Regex.Escape(kvp.Key)}\b";
                    source = Regex.Replace(source, pattern, kvp.Value, RegexOptions.IgnoreCase);
                }

                // Update the module's source code.
                mod.SourceCode = source;
            }
        }

        // 8. Save the updated document.
        loadedDoc.Save(updatedPath);

        // 9. Simple verification output (optional).
        Console.WriteLine($"Document created: {initialPath}");
        Console.WriteLine($"Document updated: {updatedPath}");
    }
}
