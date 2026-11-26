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
        // Paths for the JSON definition file and the output macro-enabled document.
        string jsonPath = "macros.json";
        string outputPath = "ImportedMacros.docm";

        // Create sample macro definitions and write them to a JSON file.
        var sampleMacros = new List<MacroDefinition>
        {
            new MacroDefinition
            {
                Name = "MacroHello",
                SourceCode = "Sub MacroHello()\n    MsgBox \"Hello from VBA!\"\nEnd Sub"
            },
            new MacroDefinition
            {
                Name = "MacroGoodbye",
                SourceCode = "Sub MacroGoodbye()\n    MsgBox \"Goodbye from VBA!\"\nEnd Sub"
            }
        };
        string jsonString = JsonSerializer.Serialize(sampleMacros, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonString);

        // Read macro definitions from the JSON file.
        string jsonContent = File.ReadAllText(jsonPath);
        List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(jsonContent);

        // Create a blank Word document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject project = new VbaProject();
        project.Name = "ImportedMacrosProject";
        doc.VbaProject = project;

        // Add each macro as a VBA module.
        if (macros != null)
        {
            foreach (var macro in macros)
            {
                // Ensure source code is not null.
                string source = macro.SourceCode ?? string.Empty;

                VbaModule module = new VbaModule();
                module.Name = macro.Name;
                module.Type = VbaModuleType.ProceduralModule;
                module.SourceCode = source;

                doc.VbaProject.Modules.Add(module);
            }
        }

        // Save the document as a macro-enabled .docm file.
        doc.Save(outputPath);

        // Load the saved document and output module information for verification.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine($"Document has macros: {loadedDoc.HasMacros}");
        if (loadedDoc.HasMacros)
        {
            foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
            {
                Console.WriteLine($"Module: {mod.Name}");
                Console.WriteLine($"Source:\n{mod.SourceCode}");
            }
        }
    }

    // Helper class matching the JSON structure for macro definitions.
    private class MacroDefinition
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }
}
