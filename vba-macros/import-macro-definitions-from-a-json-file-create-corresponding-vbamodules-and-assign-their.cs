using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

namespace AsposeWordsVbaImport
{
    // Represents a macro definition read from JSON.
    public class MacroDefinition
    {
        public string Name { get; set; }
        public string Type { get; set; }          // Expected values: "ProceduralModule", "DocumentModule", "ClassModule", "DesignerModule"
        public string SourceCode { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the JSON input and the resulting macro‑enabled document.
            string jsonPath = Path.Combine(Environment.CurrentDirectory, "macros.json");
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ImportedMacros.docm");

            // Ensure a sample JSON file exists.
            if (!File.Exists(jsonPath))
            {
                var sampleMacros = new List<MacroDefinition>
                {
                    new MacroDefinition
                    {
                        Name = "Module1",
                        Type = "ProceduralModule",
                        SourceCode = "Sub HelloWorld()\n    MsgBox \"Hello, World!\"\nEnd Sub"
                    },
                    new MacroDefinition
                    {
                        Name = "Module2",
                        Type = "ProceduralModule",
                        SourceCode = "Sub AddNumbers()\n    Dim a As Integer, b As Integer\n    a = 5\n    b = 7\n    MsgBox \"Sum = \" & (a + b)\nEnd Sub"
                    }
                };
                string json = JsonSerializer.Serialize(sampleMacros, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(jsonPath, json);
            }

            // Read macro definitions from the JSON file.
            string jsonContent = File.ReadAllText(jsonPath);
            List<MacroDefinition> macros = JsonSerializer.Deserialize<List<MacroDefinition>>(jsonContent);

            // Create a blank Word document.
            Document doc = new Document();

            // Create a new VBA project and assign it to the document.
            VbaProject project = new VbaProject();
            project.Name = "ImportedMacrosProject";
            doc.VbaProject = project;

            // Add a VbaModule for each macro definition.
            foreach (MacroDefinition macro in macros)
            {
                VbaModule module = new VbaModule();
                module.Name = macro.Name ?? "UnnamedModule";

                // Parse the module type string to the corresponding enum value.
                if (Enum.TryParse<VbaModuleType>(macro.Type, out VbaModuleType moduleType))
                    module.Type = moduleType;
                else
                    module.Type = VbaModuleType.ProceduralModule; // Default fallback.

                // Guard against null source code.
                module.SourceCode = macro.SourceCode ?? string.Empty;

                // Add the module to the VBA project.
                doc.VbaProject.Modules.Add(module);
            }

            // Save the document as a macro‑enabled .docm file.
            doc.Save(outputPath);
        }
    }
}
