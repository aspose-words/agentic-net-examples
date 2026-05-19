using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    // Simple configuration model.
    private class CopyConfig
    {
        public List<string> ModulesToCopy { get; set; } = new List<string>();
    }

    public static void Main()
    {
        // Paths for the sample files.
        string sourcePath = "source.docm";
        string destPath = "dest.docm";
        string configPath = "config.json";

        // 1. Create a configuration file that lists the module names to copy.
        var config = new CopyConfig
        {
            ModulesToCopy = new List<string> { "ModuleA", "ModuleC" }
        };
        File.WriteAllText(configPath, JsonSerializer.Serialize(config));

        // 2. Create a source macro‑enabled document with several VBA modules.
        Document sourceDoc = new Document();

        // Ensure the source document has a VBA project.
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        sourceDoc.VbaProject = sourceProject;

        // Helper to add a module.
        void AddModule(string name, string code)
        {
            VbaModule module = new VbaModule
            {
                Name = name,
                Type = VbaModuleType.ProceduralModule,
                SourceCode = code
            };
            sourceDoc.VbaProject.Modules.Add(module);
        }

        AddModule("ModuleA", "Sub MacroA()\n    MsgBox \"Hello from A\"\nEnd Sub");
        AddModule("ModuleB", "Sub MacroB()\n    MsgBox \"Hello from B\"\nEnd Sub");
        AddModule("ModuleC", "Sub MacroC()\n    MsgBox \"Hello from C\"\nEnd Sub");

        // Save the source document as a macro‑enabled file.
        sourceDoc.Save(sourcePath);

        // 3. Create an empty destination document that will receive the copied modules.
        Document destDoc = new Document();

        // Ensure the destination document has a VBA project.
        VbaProject destProject = new VbaProject { Name = "DestinationProject" };
        destDoc.VbaProject = destProject;

        // 4. Load the configuration.
        CopyConfig loadedConfig = JsonSerializer.Deserialize<CopyConfig>(File.ReadAllText(configPath));

        // 5. Load the source document (already saved) to access its VBA project.
        Document loadedSource = new Document(sourcePath);

        // 6. Copy the specified modules.
        foreach (string moduleName in loadedConfig.ModulesToCopy)
        {
            // Retrieve the module from the source; guard against missing modules.
            VbaModule sourceModule = loadedSource.VbaProject?.Modules[moduleName];
            if (sourceModule != null)
            {
                // Clone the module to avoid referencing the original object.
                VbaModule clonedModule = sourceModule.Clone();

                // Add the cloned module to the destination project.
                destDoc.VbaProject.Modules.Add(clonedModule);
            }
        }

        // 7. Save the destination document as a macro‑enabled file.
        destDoc.Save(destPath);

        // 8. Simple validation output (no user interaction).
        Console.WriteLine($"Source document modules: {loadedSource.VbaProject.Modules.Count}");
        Console.WriteLine($"Destination document modules: {destDoc.VbaProject.Modules.Count}");
    }
}
