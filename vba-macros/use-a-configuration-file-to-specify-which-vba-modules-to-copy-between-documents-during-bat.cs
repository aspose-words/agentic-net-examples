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
        // Define paths for the sample files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string sourcePath = Path.Combine(outputDir, "Source.docm");
        string destinationPath = Path.Combine(outputDir, "Destination.docm");
        string configPath = Path.Combine(outputDir, "modules.json");

        // -----------------------------------------------------------------
        // 1. Create a source macro‑enabled document with two VBA modules.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Create a new VBA project for the source document.
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        sourceDoc.VbaProject = sourceProject;

        // Module A
        VbaModule moduleA = new VbaModule
        {
            Name = "ModuleA",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloA()\n    MsgBox \"Hello from Module A\"\nEnd Sub"
        };
        sourceDoc.VbaProject.Modules.Add(moduleA);

        // Module B
        VbaModule moduleB = new VbaModule
        {
            Name = "ModuleB",
            Type = VbaModuleType.ProceduralModule,
            SourceCode = "Sub HelloB()\n    MsgBox \"Hello from Module B\"\nEnd Sub"
        };
        sourceDoc.VbaProject.Modules.Add(moduleB);

        // Save the source document as a macro‑enabled file.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Create an empty destination macro‑enabled document.
        // -----------------------------------------------------------------
        Document destDoc = new Document();
        // No VBA project yet; it will be created later if needed.
        destDoc.Save(destinationPath);

        // -----------------------------------------------------------------
        // 3. Write a simple configuration file that lists modules to copy.
        // -----------------------------------------------------------------
        var modulesToCopy = new List<string> { "ModuleA" };
        string json = JsonSerializer.Serialize(modulesToCopy, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, json);

        // -----------------------------------------------------------------
        // 4. Load the configuration.
        // -----------------------------------------------------------------
        List<string> configModules = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(configPath));

        // -----------------------------------------------------------------
        // 5. Load source and destination documents.
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(destinationPath);

        // Ensure the destination has a VBA project.
        if (dst.VbaProject == null)
        {
            VbaProject newProject = new VbaProject { Name = "DestinationProject" };
            dst.VbaProject = newProject;
        }

        // -----------------------------------------------------------------
        // 6. Copy the specified modules from source to destination.
        // -----------------------------------------------------------------
        foreach (string moduleName in configModules)
        {
            VbaModule srcModule = src.VbaProject?.Modules[moduleName];
            if (srcModule != null)
            {
                // If a module with the same name already exists in the destination, remove it.
                VbaModule existing = dst.VbaProject.Modules[moduleName];
                if (existing != null)
                {
                    dst.VbaProject.Modules.Remove(existing);
                }

                // Clone the source module and add it to the destination project.
                VbaModule cloned = srcModule.Clone();
                dst.VbaProject.Modules.Add(cloned);
            }
        }

        // -----------------------------------------------------------------
        // 7. Save the updated destination document.
        // -----------------------------------------------------------------
        string resultPath = Path.Combine(outputDir, "Result.docm");
        dst.Save(resultPath);

        // -----------------------------------------------------------------
        // 8. Simple validation: output the names of modules now present in the result.
        // -----------------------------------------------------------------
        Document resultDoc = new Document(resultPath);
        Console.WriteLine("Modules present in the resulting document:");
        foreach (VbaModule mod in resultDoc.VbaProject.Modules)
        {
            Console.WriteLine($"- {mod.Name}");
        }

        // The program finishes automatically; no user interaction required.
    }
}
