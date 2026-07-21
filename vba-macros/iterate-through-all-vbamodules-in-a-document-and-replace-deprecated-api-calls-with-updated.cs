using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "OriginalMacros.docm");
        string modifiedPath = Path.Combine(Directory.GetCurrentDirectory(), "ModifiedMacros.docm");

        // 1. Create a new blank document.
        Document doc = new Document();

        // 2. Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // 3. Add a procedural module with sample VBA code that contains deprecated API calls.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub Example()
    ' Deprecated call
    Call OldFunction()
    ' Another deprecated call
    Selection.TypeParagraph
End Sub"
        };
        doc.VbaProject.Modules.Add(module1);

        // 4. Add a second module to demonstrate handling multiple modules.
        VbaModule module2 = new VbaModule
        {
            Name = "Module2",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Function Compute()
    ' Deprecated call
    Dim result As Long
    result = OldFunction()
    Compute = result
End Function"
        };
        doc.VbaProject.Modules.Add(module2);

        // 5. Save the original document (contains the deprecated calls).
        doc.Save(originalPath);

        // 6. Iterate through all VBA modules and replace deprecated API calls.
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            // Guard against null source code.
            string source = module.SourceCode ?? string.Empty;

            // Replace deprecated calls with their updated equivalents.
            source = source.Replace("OldFunction()", "NewFunction()");
            source = source.Replace("Selection.TypeParagraph", "Selection.TypeText \"\\n\"");

            // Update the module's source code.
            module.SourceCode = source;
        }

        // 7. Save the modified document.
        doc.Save(modifiedPath);

        // 8. Output simple verification to the console.
        Console.WriteLine($"Original document saved to: {originalPath}");
        Console.WriteLine($"Modified document saved to: {modifiedPath}");
        Console.WriteLine("Replacement of deprecated API calls completed.");
    }
}
