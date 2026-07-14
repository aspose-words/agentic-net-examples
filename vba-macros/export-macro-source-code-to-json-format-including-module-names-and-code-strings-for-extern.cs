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
        // Path for the temporary macro-enabled document.
        string docPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleMacro.docm");

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject();
            project.Name = "SampleProject";
            doc.VbaProject = project;
        }

        // Add a procedural module with sample VBA code.
        VbaModule module1 = new VbaModule();
        module1.Name = "ModuleOne";
        module1.Type = VbaModuleType.ProceduralModule;
        module1.SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub
";
        doc.VbaProject.Modules.Add(module1);

        // Add a class module with sample VBA code.
        VbaModule module2 = new VbaModule();
        module2.Name = "ClassOne";
        module2.Type = VbaModuleType.ClassModule;
        module2.SourceCode = @"
Public Sub Greet()
    MsgBox ""Greetings from ClassOne!""
End Sub
";
        doc.VbaProject.Modules.Add(module2);

        // Save the document in a macro-enabled format.
        doc.Save(docPath);

        // Load the document (optional, demonstrates loading from file).
        Document loadedDoc = new Document(docPath);

        // Prepare a list to hold module information.
        List<MacroInfo> macros = new List<MacroInfo>();

        // Iterate over all VBA modules.
        foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
        {
            // Guard against null source code.
            string code = mod.SourceCode ?? string.Empty;

            macros.Add(new MacroInfo
            {
                Name = mod.Name,
                SourceCode = code
            });
        }

        // Serialize the list to JSON.
        string json = JsonSerializer.Serialize(macros, new JsonSerializerOptions { WriteIndented = true });

        // Output JSON to console.
        Console.WriteLine(json);

        // Optionally, write JSON to a file.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "macros.json");
        File.WriteAllText(jsonPath, json);
    }

    // Simple DTO for JSON serialization.
    private class MacroInfo
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }
}
