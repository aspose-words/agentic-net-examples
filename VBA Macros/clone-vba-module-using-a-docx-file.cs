using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Vba;

class CloneVbaModuleExample
{
    static void Main()
    {
        // Paths to the source document (containing the VBA module) and the destination document.
        string dataDir = @"C:\Data\";
        string sourcePath = Path.Combine(dataDir, "Source.docm");   // Document with VBA project
        string destinationPath = Path.Combine(dataDir, "ClonedModule.docm");

        // Load the source document.
        Document srcDoc = new Document(sourcePath);

        // Create a new empty destination document.
        Document destDoc = new Document();

        // Ensure the destination document has a VBA project.
        if (destDoc.VbaProject == null)
            destDoc.VbaProject = new VbaProject();

        // Get the VBA project from the source document.
        VbaProject srcVbaProject = srcDoc.VbaProject;

        // Choose the module to clone (by name). Adjust the name as needed.
        string moduleNameToClone = "Module1";
        VbaModule srcModule = srcVbaProject.Modules[moduleNameToClone];

        // Clone the selected module.
        VbaModule clonedModule = srcModule.Clone();

        // If the destination already contains a module with the same name, remove it.
        VbaModule existingModule = destDoc.VbaProject.Modules[moduleNameToClone];
        if (existingModule != null)
            destDoc.VbaProject.Modules.Remove(existingModule);

        // Add the cloned module to the destination document's VBA project.
        destDoc.VbaProject.Modules.Add(clonedModule);

        // Save the destination document. Use the DOCM format to preserve macros.
        destDoc.Save(destinationPath, SaveFormat.Docm);
    }
}
