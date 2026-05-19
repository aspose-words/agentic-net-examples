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
        public string Code { get; set; }
    }

    public static void Main()
    {
        // Path for the temporary JSON file containing macro definitions.
        string jsonPath = Path.Combine(Environment.CurrentDirectory, "macros.json");

        // Sample JSON content: an array of macro definitions.
        string sampleJson = @"
        [
            {
                ""Name"": ""HelloMacro"",
                ""Code"": ""Sub HelloMacro()\n    MsgBox \""Hello from VBA!\""\nEnd Sub""
            },
            {
                ""Name"": ""AddNumbers"",
                ""Code"": ""Function AddNumbers(a As Integer, b As Integer) As Integer\n    AddNumbers = a + b\nEnd Function""
            }
        ]";

        // Write the sample JSON to the file system.
        File.WriteAllText(jsonPath, sampleJson);

        // Read and deserialize the JSON file into a list of macro definitions.
        string jsonContent = File.ReadAllText(jsonPath);
        List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(jsonContent);

        // Create a new blank Word document.
        Document doc = new Document();

        // Create a new VBA project and assign it to the document.
        VbaProject vbaProject = new VbaProject
        {
            Name = "ImportedMacrosProject"
        };
        doc.VbaProject = vbaProject;

        // Add a VBA module for each macro definition.
        foreach (MacroDefinition macro in macros)
        {
            // Guard against null source code.
            string source = macro.Code ?? string.Empty;

            VbaModule module = new VbaModule
            {
                Name = macro.Name,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = source
            };

            doc.VbaProject.Modules.Add(module);
        }

        // Save the document as a macro‑enabled file.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "Result.docm");
        doc.Save(outputPath);

        // Simple verification: output the names of the added modules.
        Console.WriteLine($"Document saved to: {outputPath}");
        Console.WriteLine("Added VBA modules:");
        foreach (VbaModule module in doc.VbaProject.Modules)
        {
            Console.WriteLine($"- {module.Name}");
        }
    }
}
