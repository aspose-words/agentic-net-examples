using System;
using Aspose.Words;
using Aspose.Words.Vba;

namespace VbaModuleCopyExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a source macro‑enabled document.
            Document sourceDoc = new Document();

            // Ensure the source document has a VBA project.
            if (sourceDoc.VbaProject == null)
                sourceDoc.VbaProject = new VbaProject();

            // Add a VBA module to the source document.
            string moduleName = "SampleModule";
            string moduleCode = @"Sub ShowMessage()
    MsgBox ""Hello from source!""
End Sub";

            // VbaModule has a parameterless constructor; set properties manually.
            VbaModule sourceModuleToAdd = new VbaModule
            {
                Name = moduleName,
                SourceCode = moduleCode
            };
            sourceDoc.VbaProject.Modules.Add(sourceModuleToAdd);

            // Save the source document in a macro‑enabled format.
            string sourcePath = "source.docm";
            sourceDoc.Save(sourcePath, SaveFormat.Docm);

            // Create a destination macro‑enabled document.
            Document destDoc = new Document();

            // Ensure the destination document has a VBA project.
            if (destDoc.VbaProject == null)
                destDoc.VbaProject = new VbaProject();

            // Load the source document to access its VBA module.
            Document loadedSource = new Document(sourcePath);
            VbaModule loadedSourceModule = loadedSource.VbaProject?.Modules[moduleName];

            // Guard against missing module.
            if (loadedSourceModule != null)
            {
                // Copy the module into the destination document.
                VbaModule copiedModule = new VbaModule
                {
                    Name = loadedSourceModule.Name,
                    SourceCode = loadedSourceModule.SourceCode ?? string.Empty
                };
                destDoc.VbaProject.Modules.Add(copiedModule);
            }

            // Save the destination document.
            string destPath = "dest.docm";
            destDoc.Save(destPath, SaveFormat.Docm);

            // Simple validation: confirm the module exists in the destination.
            bool moduleCopied = destDoc.VbaProject?.Modules[moduleName] != null;
            Console.WriteLine($"Module '{moduleName}' copied: {moduleCopied}");
        }
    }
}
