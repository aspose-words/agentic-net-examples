using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaProjectExample
{
    static void Main()
    {
        // Load the source document that contains a VBA project.
        Document srcDoc = new Document("VBA project.docm");

        // Create a new empty document that will receive the cloned VBA project.
        Document destDoc = new Document();

        // Clone the entire VBA project (includes references).
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document.
        destDoc.VbaProject = clonedProject;

        // The destination document may contain a default module (e.g., "Module1").
        // Remove all existing modules to avoid duplicates.
        while (destDoc.VbaProject.Modules.Count > 0)
        {
            destDoc.VbaProject.Modules.Remove(destDoc.VbaProject.Modules[0]);
        }

        // Add cloned modules from the source document preserving their original order.
        foreach (VbaModule srcModule in srcDoc.VbaProject.Modules)
        {
            VbaModule clonedModule = srcModule.Clone();
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the document with the duplicated VBA project, preserving module order and references.
        destDoc.Save("VbaProject.CloneVbaProject.docm");
    }
}
