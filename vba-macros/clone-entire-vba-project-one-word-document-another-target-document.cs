using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        const string sourcePath = "Source.docm";
        const string destinationPath = "Destination.docm";

        // Ensure a source .docm file exists. If not, create a minimal one.
        if (!File.Exists(sourcePath))
        {
            var tempDoc = new Document();
            // Save as .docm to allow a VBA project container (even if empty).
            tempDoc.Save(sourcePath);
        }

        // Load the source document that contains the VBA project.
        Document srcDoc = new Document(sourcePath);

        // Create a new blank document.
        Document dstDoc = new Document();

        // Clone the entire VBA project from the source document, if it exists.
        if (srcDoc.VbaProject != null)
        {
            VbaProject clonedProject = srcDoc.VbaProject.Clone();
            dstDoc.VbaProject = clonedProject;

            // Replace any modules that may already exist in the destination with the cloned ones.
            foreach (VbaModule srcModule in srcDoc.VbaProject.Modules)
            {
                // Clone the individual VBA module.
                VbaModule clonedModule = srcModule.Clone();

                // If a module with the same name already exists, remove it.
                VbaModule existing = dstDoc.VbaProject.Modules[clonedModule.Name];
                if (existing != null)
                    dstDoc.VbaProject.Modules.Remove(existing);

                // Add the cloned module to the destination VBA project.
                dstDoc.VbaProject.Modules.Add(clonedModule);
            }
        }

        // Save the destination document (it will retain the macros if any were present).
        dstDoc.Save(destinationPath);

        Console.WriteLine($"Cloned document saved to '{destinationPath}'.");
    }
}
