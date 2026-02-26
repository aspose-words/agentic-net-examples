using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModules
{
    static void Main()
    {
        // Path to the source document that contains VBA macros (must be .docm or .dotm)
        string sourcePath = @"C:\Docs\Source.docm";

        // Path where the cloned document will be saved
        string destinationPath = @"C:\Docs\Cloned.docm";

        // Load the source document
        Document srcDoc = new Document(sourcePath);

        // Create a new empty document that will receive the cloned VBA project
        Document destDoc = new Document();

        // ------------------------------------------------------------
        // Clone the entire VBA project from the source document
        // ------------------------------------------------------------
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document
        destDoc.VbaProject = clonedProject;

        // ------------------------------------------------------------
        // Optionally, clone a single module separately (e.g., "Module1")
        // ------------------------------------------------------------
        VbaModule originalModule = srcDoc.VbaProject.Modules["Module1"];
        if (originalModule != null)
        {
            // Perform a deep copy of the module
            VbaModule clonedModule = originalModule.Clone();

            // If a module with the same name already exists in the destination,
            // remove it to avoid a name conflict
            VbaModule existing = destDoc.VbaProject.Modules["Module1"];
            if (existing != null)
            {
                destDoc.VbaProject.Modules.Remove(existing);
            }

            // Add the cloned module to the destination project's collection
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document with the cloned VBA macros
        destDoc.Save(destinationPath);
    }
}
