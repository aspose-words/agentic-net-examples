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

            // Load and deserialize the configuration.
            ConfigRoot config = JsonSerializer.Deserialize<ConfigRoot>(File.ReadAllText(configPath));

            foreach (CopyInstruction instruction in config.Mappings)
            {
                // Load source and destination documents using the provided constructors.
                Document srcDoc = new Document(instruction.Source);
                Document dstDoc = new Document(instruction.Destination);

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

                // Save the modified destination document using the provided Save method.
                dstDoc.Save(instruction.Destination);
                Console.WriteLine($"Processed '{instruction.Source}' -> '{instruction.Destination}'.");
            }
        }
    }
}
