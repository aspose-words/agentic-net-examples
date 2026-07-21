using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    // Represents a macro definition read from JSON.
    private class MacroDefinition
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }

    public static void Main()
    {
        // Prepare a sample JSON file with macro definitions.
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "macros.json");
        var sampleMacros = new List<MacroDefinition>
        {
            new MacroDefinition
            {
                Name = "HelloWorldModule",
                SourceCode = @"
Sub HelloWorld()
    MsgBox ""Hello, World!""
End Sub"
            },
            new MacroDefinition
            {
                Name = "AddNumbersModule",
                SourceCode = @"
Function AddNumbers(a As Integer, b As Integer) As Integer
    AddNumbers = a + b
End Function"
            }
        };
        string jsonContent = JsonSerializer.Serialize(sampleMacros, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonContent);

        // Load macro definitions from the JSON file.
        string readJson = File.ReadAllText(jsonPath);
        List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(readJson);

        // Create a new blank Word document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        VbaProject vbaProject = new VbaProject
        {
            Name = "ImportedMacrosProject"
        };
        doc.VbaProject = vbaProject;

        // Add a VBA module for each macro definition.
        foreach (var macro in macros)
        {
            // Guard against null source code.
            string source = macro.SourceCode ?? string.Empty;

            VbaModule module = new VbaModule
            {
                Name = macro.Name,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = source
            };

            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ImportedMacros.docm");
        doc.Save(outputPath);

        // Load the saved document to verify that macros were added.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine($"Document has macros: {loadedDoc.HasMacros}");
        Console.WriteLine($"Number of VBA modules: {loadedDoc.VbaProject.Modules.Count}");

        foreach (VbaModule mod in loadedDoc.VbaProject.Modules)
        {
            Console.WriteLine($"Module: {mod.Name}");
            Console.WriteLine("Source code:");
            Console.WriteLine(mod.SourceCode);
            Console.WriteLine(new string('-', 40));
        }
    }
}
