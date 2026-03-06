using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Load the source document that contains the VBA project.
        Document srcDoc = new Document("Source.docm");

        // Create an empty destination document.
        Document destDoc = new Document();

        // Clone the entire VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();
        destDoc.VbaProject = clonedProject;

        // The cloned project may contain a default module (e.g., "Module1").
        // Remove all modules that were automatically added during cloning.
        foreach (var module in destDoc.VbaProject.Modules.ToList())
        {
            destDoc.VbaProject.Modules.Remove(module);
        }

        // Deep‑clone each module from the source document and add it to the destination.
        foreach (VbaModule srcModule in srcDoc.VbaProject.Modules)
        {
            VbaModule clonedModule = srcModule.Clone();
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document with the cloned VBA project.
        destDoc.Save("ClonedVbaProject.docm");
    }
}
