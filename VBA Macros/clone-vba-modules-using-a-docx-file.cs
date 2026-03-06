using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModules
{
    static void Main()
    {
        // Path to the source DOCX that contains VBA macros.
        string sourcePath = @"C:\Docs\Source.docx";

        // Path where the destination document will be saved.
        string destinationPath = @"C:\Docs\Cloned.docx";

        // Load the source document.
        Document srcDoc = new Document(sourcePath);

        // Create a new empty document that will receive the cloned VBA modules.
        Document dstDoc = new Document();

        // Get the VBA project from the source document.
        VbaProject srcProject = srcDoc.VbaProject;

        // If the source document contains a VBA project, clone it.
        if (srcProject != null && srcProject.Modules.Count > 0)
        {
            // Deep clone the entire VBA project (includes all modules).
            VbaProject clonedProject = srcProject.Clone();

            // Assign the cloned project to the destination document.
            dstDoc.VbaProject = clonedProject;

            // OPTIONAL: If you need to replace specific modules rather than copy the whole project,
            // you can iterate through the modules, clone each one, and add it to the destination.
            // The following block demonstrates that approach.

            /*
            foreach (VbaModule srcModule in srcProject.Modules)
            {
                // Remove any existing module in the destination with the same name.
                VbaModule existing = dstDoc.VbaProject.Modules[srcModule.Name];
                if (existing != null)
                {
                    dstDoc.VbaProject.Modules.Remove(existing);
                }

                // Clone the individual module.
                VbaModule clonedModule = srcModule.Clone();

                // Add the cloned module to the destination project's collection.
                dstDoc.VbaProject.Modules.Add(clonedModule);
            }
            */
        }

        // Save the destination document (as DOCX; macros are preserved in the .docm format if needed).
        dstDoc.Save(destinationPath, SaveFormat.Docx);
    }
}
