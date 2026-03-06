using System;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModule
{
    static void Main()
    {
        // Path to the folder that contains the documents.
        string dataDir = @"C:\Data\";

        // Load the source document that contains the VBA module to be cloned.
        Document srcDoc = new Document(dataDir + "Source.docx");

        // Create a new (empty) destination document.
        Document destDoc = new Document();

        // Ensure the destination document has a VBA project.
        if (destDoc.VbaProject == null)
        {
            destDoc.VbaProject = new VbaProject();
            destDoc.VbaProject.Name = "ClonedProject";
        }

        // Name of the VBA module to clone from the source document.
        string moduleName = "Module1";

        // Retrieve the module from the source document.
        VbaModule srcModule = srcDoc.VbaProject.Modules[moduleName];

        if (srcModule != null)
        {
            // Perform a deep clone of the module.
            VbaModule clonedModule = srcModule.Clone();

            // If a module with the same name already exists in the destination,
            // remove it to avoid a name conflict.
            VbaModule existingModule = destDoc.VbaProject.Modules[moduleName];
            if (existingModule != null)
            {
                destDoc.VbaProject.Modules.Remove(existingModule);
            }

            // Add the cloned module to the destination document's VBA project.
            destDoc.VbaProject.Modules.Add(clonedModule);
        }

        // Save the destination document (must be saved as a macro-enabled format).
        destDoc.Save(dataDir + "ClonedModule.docm");
    }
}
