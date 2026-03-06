using System;
using Aspose.Words;
using Aspose.Words.Vba;

class Program
{
    static void Main()
    {
        // Path to the folder that contains the source document.
        string dataDir = @"C:\Data\";

        // Load the source document that already contains a VBA project.
        Document srcDoc = new Document(dataDir + "Source.docm");

        // Create a new blank document that will receive the cloned VBA project.
        Document destDoc = new Document();

        // Deep clone the VBA project from the source document.
        VbaProject clonedProject = srcDoc.VbaProject.Clone();

        // Assign the cloned project to the destination document.
        destDoc.VbaProject = clonedProject;

        // The clone operation also copies all modules, so a module with the same name
        // (e.g., "Module1") already exists in the destination. Remove it to avoid a conflict.
        VbaModule existingModule = destDoc.VbaProject.Modules["Module1"];
        if (existingModule != null)
        {
            destDoc.VbaProject.Modules.Remove(existingModule);
        }

        // Clone the specific module you want to keep from the source document.
        VbaModule copiedModule = srcDoc.VbaProject.Modules["Module1"].Clone();

        // Add the cloned module to the destination document's VBA project.
        destDoc.VbaProject.Modules.Add(copiedModule);

        // Save the resulting document (must be a macro-enabled format).
        destDoc.Save(dataDir + "Cloned.docm");
    }
}
