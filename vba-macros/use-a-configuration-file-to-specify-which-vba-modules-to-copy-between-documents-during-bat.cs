using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Vba;

public class Program
{
    public static void Main()
    {
        // Define paths for sample documents and configuration file.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        string sourcePath = Path.Combine(dataDir, "Source.docm");
        string targetPath = Path.Combine(dataDir, "Target.docm");
        string configPath = Path.Combine(dataDir, "CopyModules.config");

        // -----------------------------------------------------------------
        // Step 1: Create a source macro-enabled document with several VBA modules.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();

        // Ensure the source document has a VBA project.
        VbaProject sourceProject = new VbaProject { Name = "SourceProject" };
        sourceDoc.VbaProject = sourceProject;

        // Helper to add a module.
        void AddModule(string name, string code)
        {
            VbaModule module = new VbaModule();
            module.Name = name;
            module.Type = VbaModuleType.ProceduralModule;
            module.SourceCode = code;
            sourceDoc.VbaProject.Modules.Add(module);
        }

        AddModule("ModuleA", "Sub MacroA()\n    MsgBox \"Hello from A\"\nEnd Sub");
        AddModule("ModuleB", "Sub MacroB()\n    MsgBox \"Hello from B\"\nEnd Sub");
        AddModule("ModuleC", "Sub MacroC()\n    MsgBox \"Hello from C\"\nEnd Sub");

        sourceDoc.Save(sourcePath); // Saved as .docm (macro-enabled).

        // -----------------------------------------------------------------
        // Step 2: Create an empty target macro-enabled document.
        // -----------------------------------------------------------------
        Document targetDoc = new Document();

        // Create an empty VBA project for the target.
        VbaProject targetProject = new VbaProject { Name = "TargetProject" };
        targetDoc.VbaProject = targetProject;

        targetDoc.Save(targetPath); // Saved as .docm.

        // -----------------------------------------------------------------
        // Step 3: Write a simple configuration file listing modules to copy.
        // -----------------------------------------------------------------
        // Each line contains the name of a module to copy from source to target.
        File.WriteAllLines(configPath, new[] { "ModuleA", "ModuleC" });

        // -----------------------------------------------------------------
        // Step 4: Load documents and copy specified modules according to config.
        // -----------------------------------------------------------------
        Document src = new Document(sourcePath);
        Document dst = new Document(targetPath);

        // Ensure destination has a VBA project (it should, but guard anyway).
        if (dst.VbaProject == null)
        {
            dst.VbaProject = new VbaProject { Name = "TargetProject" };
        }

        // Read module names from configuration file.
        List<string> modulesToCopy = new List<string>();
        foreach (string line in File.ReadAllLines(configPath))
        {
            string trimmed = line.Trim();
            if (!string.IsNullOrEmpty(trimmed))
                modulesToCopy.Add(trimmed);
        }

        // Copy each requested module.
        foreach (string moduleName in modulesToCopy)
        {
            VbaModule srcModule = src.VbaProject?.Modules[moduleName];
            if (srcModule != null)
            {
                // Clone the module to avoid referencing the original object.
                VbaModule cloned = srcModule.Clone();

                // If a module with the same name already exists in the target, remove it first.
                VbaModule existing = dst.VbaProject.Modules[moduleName];
                if (existing != null)
                {
                    dst.VbaProject.Modules.Remove(existing);
                }

                dst.VbaProject.Modules.Add(cloned);
            }
        }

        // Save the updated target document.
        dst.Save(targetPath);

        // -----------------------------------------------------------------
        // Step 5: Simple verification – output the names of modules now present in the target.
        // -----------------------------------------------------------------
        Console.WriteLine("Modules present in the target document after copying:");
        foreach (VbaModule module in dst.VbaProject.Modules)
        {
            Console.WriteLine($"- {module.Name}");
        }

        // The program finishes automatically.
    }
}
