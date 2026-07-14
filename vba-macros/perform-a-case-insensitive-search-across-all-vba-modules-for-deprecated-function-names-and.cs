using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    // List of deprecated function names and their replacements.
    private static readonly (string OldName, string NewName)[] DeprecatedFunctions = new[]
    {
        ("OldFunc", "NewFunc"),
        ("LegacyCalc", "ModernCalc")
    };

    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject { Name = "SampleProject" };
        doc.VbaProject = vbaProject;

        // Add a procedural module with sample VBA code containing deprecated functions.
        VbaModule module1 = new VbaModule
        {
            Name = "Module1",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub Test()
    Dim result As Long
    result = OldFunc(5)
    MsgBox result
End Sub

Function OldFunc(x As Long) As Long
    OldFunc = x * 2
End Function"
        };
        doc.VbaProject.Modules.Add(module1);

        // Add another module to demonstrate multiple modules handling.
        VbaModule module2 = new VbaModule
        {
            Name = "Helper",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Public Sub Compute()
    Dim val As Double
    val = LegacyCalc(3.14)
    Debug.Print val
End Sub

Function LegacyCalc(y As Double) As Double
    LegacyCalc = y ^ 2
End Function"
        };
        doc.VbaProject.Modules.Add(module2);

        // Save the original macro‑enabled document.
        const string originalPath = "Original.docm";
        doc.Save(originalPath);

        // Load the document back (simulating a separate operation).
        Document loadedDoc = new Document(originalPath);

        // Verify that the document indeed contains macros.
        if (!loadedDoc.HasMacros || loadedDoc.VbaProject == null)
        {
            Console.WriteLine("No VBA project found.");
            return;
        }

        // Iterate over all VBA modules and replace deprecated function names (case‑insensitive).
        foreach (VbaModule vbaModule in loadedDoc.VbaProject.Modules)
        {
            // Guard against null source code.
            string source = vbaModule.SourceCode ?? string.Empty;

            // Perform replacements for each deprecated function.
            foreach (var (oldName, newName) in DeprecatedFunctions)
            {
                // Use Regex with IgnoreCase to replace whole word occurrences.
                string pattern = $@"\b{Regex.Escape(oldName)}\b";
                source = Regex.Replace(source, pattern, newName, RegexOptions.IgnoreCase);
            }

            // Update the module's source code.
            vbaModule.SourceCode = source;
        }

        // Save the updated document.
        const string updatedPath = "Updated.docm";
        loadedDoc.Save(updatedPath);

        // Simple verification output.
        Console.WriteLine($"Original document saved as: {originalPath}");
        Console.WriteLine($"Updated document saved as: {updatedPath}");
    }
}
