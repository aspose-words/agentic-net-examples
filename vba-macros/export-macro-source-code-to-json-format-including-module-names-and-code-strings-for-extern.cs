using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define file names.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "MacroDocument.docm");
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "macros.json");

        // -----------------------------------------------------------------
        // 1. Create a new blank document.
        // -----------------------------------------------------------------
        Document doc = new Document();

        // -----------------------------------------------------------------
        // 2. Create a VBA project and assign it to the document.
        // -----------------------------------------------------------------
        VbaProject vbaProject = new VbaProject
        {
            Name = "SampleProject"
        };
        doc.VbaProject = vbaProject;

        // -----------------------------------------------------------------
        // 3. Add a couple of VBA modules with sample source code.
        // -----------------------------------------------------------------
        VbaModule module1 = new VbaModule
        {
            Name = "ModuleOne",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello from ModuleOne!""
End Sub"
        };
        doc.VbaProject.Modules.Add(module1);

        VbaModule module2 = new VbaModule
        {
            Name = "ModuleTwo",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = @"
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function"
        };
        doc.VbaProject.Modules.Add(module2);

        // -----------------------------------------------------------------
        // 4. Save the document in a macro‑enabled format.
        // -----------------------------------------------------------------
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 5. Load the document (demonstrates loading workflow).
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // -----------------------------------------------------------------
        // 6. Extract module names and source code.
        // -----------------------------------------------------------------
        var macroInfo = new List<MacroModule>();
        if (loadedDoc.HasMacros && loadedDoc.VbaProject != null)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                // Guard against null source code.
                string source = mod.SourceCode ?? string.Empty;
                macroInfo.Add(new MacroModule { Name = mod.Name, SourceCode = source });
            }
        }

        // -----------------------------------------------------------------
        // 7. Serialize the information to JSON.
        // -----------------------------------------------------------------
        var jsonOptions = new JsonSerializerOptions { WriteIndented = true };
        string json = JsonSerializer.Serialize(macroInfo, jsonOptions);

        // Write JSON to a file.
        File.WriteAllText(jsonPath, json);

        // Also output to console for visibility.
        Console.WriteLine("Exported macro information:");
        Console.WriteLine(json);
    }

    // Helper class representing a VBA module for JSON serialization.
    private class MacroModule
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }
}
