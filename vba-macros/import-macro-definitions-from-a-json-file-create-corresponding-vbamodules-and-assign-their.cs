using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class MacroDefinition
{
    public string Name { get; set; }
    public string SourceCode { get; set; }
    public string Type { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Path for the JSON file that contains macro definitions.
        const string jsonFilePath = "macros.json";

        // If the JSON file does not exist, create a sample one.
        if (!File.Exists(jsonFilePath))
        {
            var sampleMacros = new[]
            {
                new
                {
                    Name = "Module1",
                    SourceCode = "Sub Hello()\n    MsgBox \"Hello from Module1\"\nEnd Sub",
                    Type = "ProceduralModule"
                },
                new
                {
                    Name = "Class1",
                    SourceCode = "Public Sub Test()\n    ' Sample class method\nEnd Sub",
                    Type = "ClassModule"
                }
            };

            string sampleJson = JsonSerializer.Serialize(sampleMacros, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(jsonFilePath, sampleJson);
        }

        // Read and deserialize the macro definitions.
        string jsonContent = File.ReadAllText(jsonFilePath);
        List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(jsonContent);

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            VbaProject project = new VbaProject
            {
                Name = "ImportedProject"
            };
            doc.VbaProject = project;
        }

        // Add each macro as a VBA module.
        foreach (MacroDefinition macro in macros)
        {
            VbaModule module = new VbaModule
            {
                Name = macro.Name ?? "UnnamedModule",
                SourceCode = macro.SourceCode ?? string.Empty,
                Type = Enum.TryParse<VbaModuleType>(macro.Type, out var parsedType)
                    ? parsedType
                    : VbaModuleType.ProceduralModule
            };

            doc.VbaProject.Modules.Add(module);
        }

        // Save the document in a macro‑enabled format.
        const string outputPath = "Output.docm";
        doc.Save(outputPath);

        // Load the saved document to verify that modules were added.
        Document loadedDoc = new Document(outputPath);
        Console.WriteLine($"Document has macros: {loadedDoc.HasMacros}");
        Console.WriteLine($"Number of VBA modules: {loadedDoc.VbaProject?.Modules?.Count ?? 0}");

        // Iterate safely even when the collection is null.
        foreach (VbaModule mod in loadedDoc.VbaProject?.Modules ?? Enumerable.Empty<VbaModule>())
        {
            Console.WriteLine($"Module Name: {mod.Name}");
            Console.WriteLine($"Module Type: {mod.Type}");
            Console.WriteLine("Source Code:");
            Console.WriteLine(mod.SourceCode);
            Console.WriteLine(new string('-', 40));
        }
    }
}
