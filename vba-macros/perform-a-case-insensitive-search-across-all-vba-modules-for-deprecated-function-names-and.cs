using System;
using System.Collections.Generic;
using System.IO;
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
        VbaProject vbaProject = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = vbaProject;

        // Add sample VBA modules containing deprecated function names.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Sub OldFunc()
    MsgBox ""Hello from OldFunc""
End Sub"
        };
        vbaProject.Modules.Add(module1);

        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"Function DeprecatedFunction(arg As String) As String
    DeprecatedFunction = arg
End Function"
        };
        vbaProject.Modules.Add(module2);

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docm");
        doc.Save(originalPath);

        // Define deprecated function names and their replacements.
        var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "OldFunc", "NewFunc" },
            { "DeprecatedFunction", "UpdatedFunction" }
        };

        // Perform case‑insensitive search and replace in each VBA module.
        foreach (VbaModule module in vbaProject.Modules)
        {
            string source = module.SourceCode ?? string.Empty;

            foreach (var kvp in replacements)
            {
                string pattern = @"\b" + Regex.Escape(kvp.Key) + @"\b";
                source = Regex.Replace(source, pattern, kvp.Value, RegexOptions.IgnoreCase);
            }

            module.SourceCode = source;
        }

        // Save the modified document.
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "Modified.docm");
        doc.Save(modifiedPath);

        // Output paths to confirm execution.
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Modified document saved to: {modifiedPath}");
    }
}
