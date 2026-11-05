using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    // Model matching the JSON structure.
    private class MacroDefinition
    {
        public string Name { get; set; }
        public string SourceCode { get; set; }
    }

    public static void Main()
    {
        // Path for the temporary JSON file.
        const string jsonPath = "macros.json";

        // Sample macro definitions to be written to the JSON file.
        var sampleMacros = new List<MacroDefinition>
        {
            new MacroDefinition
            {
                Name = "Module1",
                SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello from VBA!\"\nEnd Sub"
            },
            new MacroDefinition
            {
                Name = "Module2",
                SourceCode = "Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function"
            }
        };

        // Serialize the sample macros to JSON and write to file.
        string jsonContent = JsonSerializer.Serialize(sampleMacros, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonContent);

        // Read the JSON file and deserialize into macro definitions.
        string readJson = File.ReadAllText(jsonPath);
        List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(readJson);

        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has a VBA project.
        if (doc.VbaProject == null)
        {
            doc.VbaProject = new VbaProject();
        }

        // Add each macro as a VBA module.
        foreach (var macro in macros)
        {
            // Guard against null source code.
            string source = macro.SourceCode ?? string.Empty;

            // Create a new VBA module and set its properties.
            VbaModule vbaModule = new VbaModule();
            vbaModule.Name = macro.Name;
            vbaModule.SourceCode = source;

            // Add the module to the VBA project.
            doc.VbaProject.Modules.Add(vbaModule);
        }

        // Save the document in a macro‑enabled format.
        const string outputPath = "output.docm";
        doc.Save(outputPath, SaveFormat.Docm);
    }
}
