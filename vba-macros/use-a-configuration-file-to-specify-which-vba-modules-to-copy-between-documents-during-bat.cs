using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    // Model that matches the JSON configuration file.
    private class Config
    {
        public List<string> ModulesToCopy { get; set; } = new List<string>();
    }

    public static void Main()
    {
        // Paths for the sample files.
        string sourcePath = Path.Combine(Directory.GetCurrentDirectory(), "source.docm");
        string destinationPath = Path.Combine(Directory.GetCurrentDirectory(), "destination.docm");
        string configPath = Path.Combine(Directory.GetCurrentDirectory(), "config.json");

        // -----------------------------------------------------------------
        // Step 1: Create a sample source document with a VBA project.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        sourceDoc.VbaProject = sourceProject;

        // Add a few VBA modules.
        CreateVbaModule(sourceProject, "Module1", "Sub Hello()\n    MsgBox \"Hello from Module1\"\nEnd Sub");
        CreateVbaModule(sourceProject, "Module2", "Sub Goodbye()\n    MsgBox \"Goodbye from Module2\"\nEnd Sub");
        CreateVbaModule(sourceProject, "ExtraModule", "Sub Extra()\n    MsgBox \"Extra module\"\nEnd Sub");

        // Save the source document as a macro‑enabled file.
        sourceDoc.Save(sourcePath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Step 2: Write a configuration file that lists modules to copy.
        // -----------------------------------------------------------------
        var config = new Config
        {
            ModulesToCopy = new List<string> { "Module1", "Module2" }
        };
        string json = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, json);

        // -----------------------------------------------------------------
        // Step 3: Load the configuration.
        // -----------------------------------------------------------------
        Config loadedConfig = JsonSerializer.Deserialize<Config>(File.ReadAllText(configPath)) ?? new Config();

        // -----------------------------------------------------------------
        // Step 4: Load the source and destination documents.
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(); // start with a blank document.

        // Ensure the destination has a VBA project.
        if (dst.VbaProject == null)
        {
            VbaProject destProject = new VbaProject { Name = "DestinationProject" };
            dst.VbaProject = destProject;
        }

        // -----------------------------------------------------------------
        // Step 5: Copy specified modules from source to destination.
        // -----------------------------------------------------------------
        foreach (string moduleName in loadedConfig.ModulesToCopy)
        {
            // Retrieve the module from the source project; may be null.
            VbaModule srcModule = src.VbaProject?.Modules[moduleName];
            if (srcModule == null)
            {
                Console.WriteLine($"Source module \"{moduleName}\" not found; skipping.");
                continue;
            }

            // Clone the module to avoid referencing the original object.
            VbaModule clonedModule = srcModule.Clone();

            // If a module with the same name already exists in the destination, remove it.
            VbaModule existing = dst.VbaProject?.Modules[moduleName];
            if (existing != null)
                dst.VbaProject?.Modules.Remove(existing);

            // Add the cloned module to the destination project.
            dst.VbaProject?.Modules.Add(clonedModule);
            Console.WriteLine($"Copied module \"{moduleName}\" to destination.");
        }

        // -----------------------------------------------------------------
        // Step 6: Save the destination document as a macro‑enabled file.
        // -----------------------------------------------------------------
        dst.Save(destinationPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // Step 7: Simple validation output.
        // -----------------------------------------------------------------
        int moduleCount = dst.VbaProject?.Modules.Count ?? 0;
        Console.WriteLine($"Destination document contains {moduleCount} VBA modules.");
        foreach (VbaModule mod in dst.VbaProject?.Modules ?? Enumerable.Empty<VbaModule>())
        {
            Console.WriteLine($"- {mod.Name}");
        }
    }

    // Helper method to create and add a VBA module to a project.
    private static void CreateVbaModule(VbaProject project, string name, string sourceCode)
    {
        VbaModule module = new VbaModule
        {
            Name = name,
            Type = VbaModuleType.ProceduralModule,
            SourceCode = sourceCode ?? string.Empty
        };
        project.Modules.Add(module);
    }
}
