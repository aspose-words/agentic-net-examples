using System;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for source, target documents and configuration file.
        string sourcePath = Path.Combine(outputDir, "source.docm");
        string targetPath = Path.Combine(outputDir, "target.docm");
        string configPath = Path.Combine(outputDir, "config.json");

        // -----------------------------------------------------------------
        // Step 1: Create a source macro-enabled document with several VBA modules.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Ensure the document has a VBA project.
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        sourceDoc.VbaProject = sourceProject;

        // Add three procedural modules with simple macros.
        for (int i = 1; i <= 3; i++)
        {
            VbaModule module = new VbaModule
            {
                Name = $"Module{i}",
                Type = VbaModuleType.ProceduralModule,
                SourceCode = $"Sub Macro{i}()\n    MsgBox \"Hello from Module{i}\"\nEnd Sub"
            };
            sourceProject.Modules.Add(module);
        }

        // Save the source document in macro-enabled format.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Step 2: Write a configuration file that lists which modules to copy.
        // The file contains a JSON array of module names.
        // -----------------------------------------------------------------
        string[] modulesToCopy = new[] { "Module1", "Module3" };
        string jsonConfig = JsonSerializer.Serialize(modulesToCopy, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, jsonConfig);

        // -----------------------------------------------------------------
        // Step 3: Load the configuration.
        // -----------------------------------------------------------------
        string configContent = File.ReadAllText(configPath);
        string[] moduleNames = JsonSerializer.Deserialize<string[]>(configContent);

        // -----------------------------------------------------------------
        // Step 4: Create the target document and ensure it has a VBA project.
        // -----------------------------------------------------------------
        Document targetDoc = new Document();

        // If the target document does not already contain a VBA project, create one.
        if (targetDoc.VbaProject == null)
        {
            VbaProject targetProject = new VbaProject { Name = "TargetProject" };
            targetDoc.VbaProject = targetProject;
        }

        // -----------------------------------------------------------------
        // Step 5: Copy the specified modules from source to target.
        // -----------------------------------------------------------------
        foreach (string moduleName in moduleNames)
        {
            // Retrieve the module from the source document; guard against missing modules.
            VbaModule sourceModule = sourceDoc.VbaProject?.Modules[moduleName];
            if (sourceModule != null)
            {
                // Clone the module to preserve its content and add it to the target document.
                VbaModule clonedModule = sourceModule.Clone();
                targetDoc.VbaProject.Modules.Add(clonedModule);
            }
        }

        // Save the target document in macro-enabled format.
        targetDoc.Save(targetPath);

        // -----------------------------------------------------------------
        // Step 6: Simple validation – list the modules present in the target document.
        // -----------------------------------------------------------------
        Console.WriteLine("Modules present in the target document:");
        foreach (VbaModule module in targetDoc.VbaProject.Modules)
        {
            Console.WriteLine($"- {module.Name}");
        }

        // The program finishes automatically; no user interaction required.
    }
}
