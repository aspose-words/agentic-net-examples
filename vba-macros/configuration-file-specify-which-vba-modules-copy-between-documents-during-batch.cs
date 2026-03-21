using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleBatchCopy
{
    // Represents a single copy instruction read from the configuration file.
    public class CopyInstruction
    {
        public string Source { get; set; }          // Path to the source document (must be macro‑enabled, e.g., .docm)
        public string Destination { get; set; }     // Path to the destination document (will be overwritten)
        public List<string> Modules { get; set; }   // Names of VBA modules to copy from source to destination
    }

    // Root object for the JSON configuration.
    public class ConfigRoot
    {
        public List<CopyInstruction> Mappings { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the JSON configuration file.
            const string configPath = "copyConfig.json";

            if (!File.Exists(configPath))
            {
                Console.WriteLine($"Configuration file '{configPath}' not found. No work to do.");
                return;
            }

            ConfigRoot config;
            try
            {
                config = JsonSerializer.Deserialize<ConfigRoot>(File.ReadAllText(configPath));
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read or deserialize configuration: {ex.Message}");
                return;
            }

            if (config?.Mappings == null || config.Mappings.Count == 0)
            {
                Console.WriteLine("No mappings defined in the configuration.");
                return;
            }

            foreach (CopyInstruction instruction in config.Mappings)
            {
                if (string.IsNullOrWhiteSpace(instruction.Source) ||
                    string.IsNullOrWhiteSpace(instruction.Destination) ||
                    instruction.Modules == null || instruction.Modules.Count == 0)
                {
                    Console.WriteLine("Invalid instruction detected; skipping.");
                    continue;
                }

                if (!File.Exists(instruction.Source))
                {
                    Console.WriteLine($"Source file '{instruction.Source}' does not exist; skipping.");
                    continue;
                }

                // Load source and destination documents.
                Document srcDoc;
                Document dstDoc;
                try
                {
                    srcDoc = new Document(instruction.Source);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load source document '{instruction.Source}': {ex.Message}");
                    continue;
                }

                try
                {
                    // If the destination does not exist, create an empty document.
                    dstDoc = File.Exists(instruction.Destination) ? new Document(instruction.Destination) : new Document();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load/create destination document '{instruction.Destination}': {ex.Message}");
                    continue;
                }

                // Ensure both documents have a VBA project; create one for the destination if missing.
                if (srcDoc.VbaProject == null)
                {
                    Console.WriteLine($"Source document '{instruction.Source}' has no VBA project. Skipping.");
                    continue;
                }

                if (dstDoc.VbaProject == null)
                {
                    dstDoc.VbaProject = new VbaProject { Name = Path.GetFileNameWithoutExtension(instruction.Destination) };
                }

                VbaModuleCollection srcModules = srcDoc.VbaProject.Modules;
                VbaModuleCollection dstModules = dstDoc.VbaProject.Modules;

                foreach (string moduleName in instruction.Modules)
                {
                    // Retrieve the module from the source by name.
                    VbaModule srcModule = srcModules[moduleName];
                    if (srcModule == null)
                    {
                        Console.WriteLine($"Module '{moduleName}' not found in source '{instruction.Source}'.");
                        continue;
                    }

                    // Clone the source module to obtain a deep copy.
                    VbaModule clonedModule = srcModule.Clone();

                    // If the destination already contains a module with the same name, remove it.
                    VbaModule existingDstModule = dstModules[moduleName];
                    if (existingDstModule != null)
                    {
                        dstModules.Remove(existingDstModule);
                    }

                    // Add the cloned module to the destination's VBA project.
                    dstModules.Add(clonedModule);
                }

                try
                {
                    dstDoc.Save(instruction.Destination);
                    Console.WriteLine($"Processed '{instruction.Source}' -> '{instruction.Destination}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save destination document '{instruction.Destination}': {ex.Message}");
                }
            }
        }
    }
}
