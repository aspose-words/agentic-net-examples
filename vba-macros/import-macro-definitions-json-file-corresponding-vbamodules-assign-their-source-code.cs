using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

namespace MacroImportExample
{
    // Represents a macro definition read from JSON.
    public class MacroDefinition
    {
        public string Name { get; set; }
        public string Type { get; set; }          // Expected values: DocumentModule, ProceduralModule, ClassModule, DesignerModule
        public string SourceCode { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the JSON file containing macro definitions.
            const string jsonPath = "macros.json";

            // Load macro definitions from the JSON file (or use a fallback if the file is missing).
            List<MacroDefinition> macros = LoadMacroDefinitions(jsonPath);

            // Create a new blank document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject vbaProject = new VbaProject
            {
                Name = "ImportedMacros"
            };
            doc.VbaProject = vbaProject;

            // Iterate over each macro definition and add it as a VbaModule.
            foreach (MacroDefinition macro in macros)
            {
                // Create a new module.
                VbaModule module = new VbaModule
                {
                    Name = macro.Name,
                    // Convert the string representation of the module type to the enum value.
                    Type = (VbaModuleType)Enum.Parse(typeof(VbaModuleType), macro.Type, ignoreCase: true),
                    SourceCode = macro.SourceCode
                };

                // Add the module to the VBA project.
                doc.VbaProject.Modules.Add(module);
            }

            // Save the document as a macro‑enabled file.
            doc.Save("ImportedMacros.docm");
        }

        // Reads the JSON file and deserializes it into a list of MacroDefinition objects.
        // If the file does not exist, returns a default list with a simple example macro.
        private static List<MacroDefinition> LoadMacroDefinitions(string path)
        {
            if (File.Exists(path))
            {
                using FileStream stream = File.OpenRead(path);
                return JsonSerializer.Deserialize<List<MacroDefinition>>(stream);
            }

            // Fallback JSON with a single example macro.
            const string fallbackJson = @"[
                {
                    ""Name"": ""ExampleModule"",
                    ""Type"": ""DocumentModule"",
                    ""SourceCode"": ""Sub HelloWorld()\n    MsgBox \""Hello, World!\""\nEnd Sub""
                }
            ]";

            return JsonSerializer.Deserialize<List<MacroDefinition>>(fallbackJson);
        }
    }
}
